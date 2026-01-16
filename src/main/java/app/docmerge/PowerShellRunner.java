// PowerShellRunner.java
package app.docmerge;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.TimeUnit;

public class PowerShellRunner {
    private final Optional<String> executable;

    public PowerShellRunner() {
        this.executable = resolveExecutable();
    }

    public Optional<String> getExecutable() {
        return executable;
    }

    public Result runScript(String script, Duration timeout) throws PowerShellExecutionException {
        if (executable.isEmpty()) {
            throw new PowerShellExecutionException("未检测到 PowerShell", new Result(-1, "", "未检测到 PowerShell"));
        }

        Path scriptFile = null;
        Process process = null;

        try {
            scriptFile = Files.createTempFile("doc-merge-com-", ".ps1");

            // 关键：powershell.exe(5.1) 读取 UTF-8 无 BOM 容易按 ANSI 解析，中文路径/中文字符串会直接炸
            writeUtf8Bom(scriptFile, script);

            List<String> command = new ArrayList<>();
            command.add(executable.get());
            command.add("-NoLogo");
            command.add("-NoProfile");
            command.add("-NonInteractive");
            command.add("-ExecutionPolicy");
            command.add("Bypass");
            command.add("-STA");
            command.add("-File");
            command.add(scriptFile.toString());

            ProcessBuilder builder = new ProcessBuilder(command);
            process = builder.start();

            boolean finished = process.waitFor(timeout.toMillis(), TimeUnit.MILLISECONDS);
            if (!finished) {
                process.destroyForcibly();
                throw new PowerShellExecutionException("PowerShell 执行超时", new Result(-1, "", "PowerShell 执行超时"));
            }

            String stdout = readStream(process.getInputStream());
            String stderr = readStream(process.getErrorStream());
            int exitCode = process.exitValue();

            Result result = new Result(exitCode, stdout, stderr);
            if (exitCode != 0) {
                throw new PowerShellExecutionException("PowerShell 执行失败", result);
            }
            return result;

        } catch (IOException e) {
            throw new PowerShellExecutionException("PowerShell 执行失败：" + e.getMessage(),
                    new Result(-1, "", e.getMessage()));
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new PowerShellExecutionException("PowerShell 执行被中断",
                    new Result(-1, "", "PowerShell 执行被中断"));
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

    private Optional<String> resolveExecutable() {
        // COM 自动化优先用 Windows PowerShell（powershell.exe），其次 pwsh
        List<String> candidates = List.of("powershell", "pwsh");
        for (String candidate : candidates) {
            try {
                Process process = new ProcessBuilder(candidate, "-NoLogo", "-NoProfile", "-NonInteractive",
                        "-Command", "$PSVersionTable.PSVersion")
                        .start();
                if (process.waitFor(5, TimeUnit.SECONDS) && process.exitValue() == 0) {
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

    private static void writeUtf8Bom(Path file, String content) throws IOException {
        byte[] bom = new byte[] {(byte) 0xEF, (byte) 0xBB, (byte) 0xBF};
        byte[] body = content.getBytes(StandardCharsets.UTF_8);
        byte[] all = new byte[bom.length + body.length];
        System.arraycopy(bom, 0, all, 0, bom.length);
        System.arraycopy(body, 0, all, bom.length, body.length);
        Files.write(file, all);
    }

    private String readStream(InputStream stream) throws IOException {
        if (stream == null) {
            return "";
        }
        return new String(stream.readAllBytes(), StandardCharsets.UTF_8);
    }

    public record Result(int exitCode, String stdout, String stderr) {
    }

    public static class PowerShellExecutionException extends IOException {
        private final Result result;

        public PowerShellExecutionException(String message, Result result) {
            super(message);
            this.result = result;
        }

        public Result getResult() {
            return result;
        }
    }
}
