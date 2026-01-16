package app.docmerge;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.concurrent.locks.ReentrantLock;

public class WordDocConverter {
    private static final int WORD_FORMAT_DOCX = 16;
    private static final ReentrantLock WORD_LOCK = new ReentrantLock(true);
    private static final Duration PROBE_TIMEOUT = Duration.ofSeconds(10);

    private Boolean supportedCache;
    private String supportMessage;

    public boolean isSupported() {
        if (supportedCache != null) {
            return supportedCache;
        }
        if (!isWindows()) {
            supportedCache = false;
            supportMessage = "仅支持 Windows + Microsoft Word";
            return false;
        }
        Optional<String> powerShell = findPowerShell();
        if (powerShell.isEmpty()) {
            supportedCache = false;
            supportMessage = "未检测到 PowerShell";
            return false;
        }
        try {
            runProbe(powerShell.get());
            supportedCache = true;
            supportMessage = "可用";
            return true;
        } catch (WordConversionException e) {
            supportedCache = false;
            supportMessage = e.getMessage();
            return false;
        }
    }

    public String getSupportMessage() {
        if (supportedCache == null) {
            isSupported();
        }
        return supportMessage == null ? "" : supportMessage;
    }

    public Path convert(Path doc) throws IOException, WordConversionException {
        List<Path> outputs = convertBatch(List.of(doc));
        if (outputs.isEmpty()) {
            throw new WordConversionException("Word 转换未生成输出文件", doc.toString(), "", "", -1);
        }
        return outputs.get(0);
    }

    public List<Path> convertBatch(List<Path> docs) throws IOException, WordConversionException {
        ConversionBatch batch = convertBatchWithTempDir(docs);
        return batch.outputs();
    }

    public ConversionBatch convertBatchWithTempDir(List<Path> docs) throws IOException, WordConversionException {
        if (docs == null || docs.isEmpty()) {
            return new ConversionBatch(List.of(), List.of(), null);
        }
        if (!isSupported()) {
            throw new WordConversionException("当前环境无法使用 Microsoft Word 进行 .doc 转换：" + getSupportMessage(),
                    docs.get(0).toString(), "", "", -1);
        }
        Optional<String> powerShell = findPowerShell();
        if (powerShell.isEmpty()) {
            throw new WordConversionException("未检测到 PowerShell，无法进行 .doc 转换", docs.get(0).toString(), "", "", -1);
        }
        Path tempDir = Files.createTempDirectory("doc-merge-word-");
        List<Path> inputs = new ArrayList<>(docs);
        List<Path> outputs = new ArrayList<>();
        for (int i = 0; i < inputs.size(); i++) {
            Path input = inputs.get(i);
            String baseName = stripExtension(input.getFileName().toString());
            String outputName = String.format("%03d_%s.docx", i + 1, baseName);
            outputs.add(tempDir.resolve(outputName));
        }

        WORD_LOCK.lock();
        try {
            runConversion(powerShell.get(), inputs, outputs, tempDir);
            for (int i = 0; i < outputs.size(); i++) {
                if (!Files.exists(outputs.get(i))) {
                    throw new WordConversionException("转换失败，未生成输出文件：" + outputs.get(i),
                            inputs.get(i).toString(), "", "", -1);
                }
            }
            return new ConversionBatch(inputs, outputs, tempDir);
        } catch (WordConversionException | IOException e) {
            cleanupConversion(new ConversionBatch(inputs, outputs, tempDir));
            throw e;
        } finally {
            WORD_LOCK.unlock();
        }
    }

    public void cleanupConversion(ConversionBatch batch) {
        if (batch == null || batch.tempDir() == null) {
            return;
        }
        try {
            Files.walkFileTree(batch.tempDir(), new SimpleFileVisitor<>() {
                @Override
                public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {
                    Files.deleteIfExists(file);
                    return FileVisitResult.CONTINUE;
                }

                @Override
                public FileVisitResult postVisitDirectory(Path dir, IOException exc) throws IOException {
                    Files.deleteIfExists(dir);
                    return FileVisitResult.CONTINUE;
                }
            });
        } catch (IOException ignored) {
            // 忽略清理失败
        }
    }

    private void runProbe(String powerShell) throws WordConversionException {
        String script = """
                $ErrorActionPreference = 'Stop'
                [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
                $OutputEncoding = [System.Text.Encoding]::UTF8
                $word = New-Object -ComObject Word.Application
                $word.Visible = $false
                $word.DisplayAlerts = 0
                try {
                    # 探测可用性
                } finally {
                    if ($word -ne $null) {
                        $word.Quit() | Out-Null
                        [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($word) | Out-Null
                    }
                }
                """;
        ProcessResult result = runPowerShell(powerShell, script, PROBE_TIMEOUT);
        if (result.exitCode() != 0) {
            throw new WordConversionException("Microsoft Word COM 探测失败",
                    "", result.stdout(), result.stderr(), result.exitCode());
        }
    }

    private void runConversion(String powerShell,
                               List<Path> inputs,
                               List<Path> outputs,
                               Path tempDir) throws IOException, WordConversionException {
        String script = buildConversionScript(inputs, outputs, tempDir);
        ProcessResult result = runPowerShell(powerShell, script, Duration.ofMinutes(10));
        if (result.exitCode() != 0) {
            String failed = inputs.isEmpty() ? "" : inputs.get(0).toString();
            throw new WordConversionException("Word 转换执行失败",
                    failed, result.stdout(), result.stderr(), result.exitCode());
        }
    }

    private String buildConversionScript(List<Path> inputs, List<Path> outputs, Path tempDir) {
        StringBuilder builder = new StringBuilder();
        builder.append("$ErrorActionPreference = 'Stop'").append(System.lineSeparator());
        builder.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("New-Item -ItemType Directory -Force -Path '")
                .append(escapePowerShell(tempDir.toString()))
                .append("' | Out-Null").append(System.lineSeparator());
        builder.append("$word = New-Object -ComObject Word.Application").append(System.lineSeparator());
        builder.append("$word.Visible = $false").append(System.lineSeparator());
        builder.append("$word.DisplayAlerts = 0").append(System.lineSeparator());
        builder.append("try {").append(System.lineSeparator());
        builder.append("  $items = @(").append(System.lineSeparator());
        for (int i = 0; i < inputs.size(); i++) {
            String input = escapePowerShell(inputs.get(i).toString());
            String output = escapePowerShell(outputs.get(i).toString());
            builder.append("    @{input='").append(input).append("'; output='").append(output).append("'}");
            if (i < inputs.size() - 1) {
                builder.append(",");
            }
            builder.append(System.lineSeparator());
        }
        builder.append("  )").append(System.lineSeparator());
        builder.append("  foreach ($item in $items) {").append(System.lineSeparator());
        builder.append("    $doc = $null").append(System.lineSeparator());
        builder.append("    try {").append(System.lineSeparator());
        builder.append("      $doc = $word.Documents.Open($item.input, $false, $true)").append(System.lineSeparator());
        builder.append("      $doc.SaveAs([ref]$item.output, [ref]")
                .append(WORD_FORMAT_DOCX).append(")").append(System.lineSeparator());
        builder.append("    } finally {").append(System.lineSeparator());
        builder.append("      if ($doc -ne $null) {").append(System.lineSeparator());
        builder.append("        $doc.Close([ref]$false) | Out-Null").append(System.lineSeparator());
        builder.append("        [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($doc) | Out-Null").append(System.lineSeparator());
        builder.append("      }").append(System.lineSeparator());
        builder.append("    }").append(System.lineSeparator());
        builder.append("  }").append(System.lineSeparator());
        builder.append("} finally {").append(System.lineSeparator());
        builder.append("  if ($word -ne $null) {").append(System.lineSeparator());
        builder.append("    $word.Quit() | Out-Null").append(System.lineSeparator());
        builder.append("    [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($word) | Out-Null").append(System.lineSeparator());
        builder.append("  }").append(System.lineSeparator());
        builder.append("}").append(System.lineSeparator());
        return builder.toString();
    }

    private ProcessResult runPowerShell(String powerShell, String script, Duration timeout) throws WordConversionException {
        Path scriptFile = null;
        try {
            scriptFile = Files.createTempFile("doc-merge-word-", ".ps1");
            Files.writeString(scriptFile, script, StandardCharsets.UTF_8);
        } catch (IOException e) {
            throw new WordConversionException("无法创建 PowerShell 脚本文件", "", "", "", -1);
        }
        Process process = null;
        try {
            List<String> command = new ArrayList<>();
            command.add(powerShell);
            command.add("-NoProfile");
            command.add("-NonInteractive");
            command.add("-ExecutionPolicy");
            command.add("Bypass");
            command.add("-STA");
            command.add("-File");
            command.add(scriptFile.toString());
            ProcessBuilder builder = new ProcessBuilder(command);
            process = builder.start();
            boolean finished = process.waitFor(timeout.toMillis(), java.util.concurrent.TimeUnit.MILLISECONDS);
            if (!finished) {
                process.destroyForcibly();
                throw new WordConversionException("PowerShell 执行超时", "", "", "", -1);
            }
            String stdout = readStream(process.getInputStream());
            String stderr = readStream(process.getErrorStream());
            return new ProcessResult(process.exitValue(), stdout, stderr);
        } catch (IOException e) {
            throw new WordConversionException("PowerShell 执行失败：" + e.getMessage(), "", "", "", -1);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new WordConversionException("PowerShell 执行被中断", "", "", "", -1);
        } finally {
            if (scriptFile != null) {
                try {
                    Files.deleteIfExists(scriptFile);
                } catch (IOException ignored) {
                    // ignore
                }
            }
            if (process != null) {
                process.destroy();
            }
        }
    }

    private String readStream(InputStream stream) throws IOException {
        if (stream == null) {
            return "";
        }
        return new String(stream.readAllBytes(), StandardCharsets.UTF_8);
    }

    private Optional<String> findPowerShell() {
        if (!isWindows()) {
            return Optional.empty();
        }
        List<String> candidates = List.of("pwsh", "powershell");
        for (String candidate : candidates) {
            try {
                Process process = new ProcessBuilder(candidate, "-NoProfile", "-NonInteractive", "-Command", "$PSVersionTable.PSVersion")
                        .start();
                if (process.waitFor(5, java.util.concurrent.TimeUnit.SECONDS) && process.exitValue() == 0) {
                    return Optional.of(candidate);
                }
            } catch (IOException | InterruptedException ignored) {
                if (ignored instanceof InterruptedException) {
                    Thread.currentThread().interrupt();
                }
            }
        }
        return Optional.empty();
    }

    private boolean isWindows() {
        return System.getProperty("os.name").toLowerCase(Locale.ROOT).contains("win");
    }

    private String stripExtension(String name) {
        int idx = name.lastIndexOf('.');
        return idx > 0 ? name.substring(0, idx) : name;
    }

    private String escapePowerShell(String value) {
        return value.replace("'", "''");
    }

    public record ConversionBatch(List<Path> inputs, List<Path> outputs, Path tempDir) {
    }

    private record ProcessResult(int exitCode, String stdout, String stderr) {
    }

    public static class WordConversionException extends IOException {
        private final String failedInput;
        private final String stdout;
        private final String stderr;
        private final int exitCode;

        public WordConversionException(String message, String failedInput, String stdout, String stderr, int exitCode) {
            super(message);
            this.failedInput = failedInput;
            this.stdout = stdout;
            this.stderr = stderr;
            this.exitCode = exitCode;
        }

        public String getFailedInput() {
            return failedInput;
        }

        public String getStdout() {
            return stdout;
        }

        public String getStderr() {
            return stderr;
        }

        public int getExitCode() {
            return exitCode;
        }
    }
}
