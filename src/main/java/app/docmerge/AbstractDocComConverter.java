package app.docmerge;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.UUID;
import java.util.concurrent.Semaphore;

public abstract class AbstractDocComConverter implements DocComConverter {
    private static final Semaphore COM_LOCK = new Semaphore(1, true);
    private static final Duration PROBE_TIMEOUT = Duration.ofSeconds(10);
    private static final Duration CONVERT_TIMEOUT = Duration.ofMinutes(10);

    private final PowerShellRunner powerShellRunner;
    private DocComProbeResult lastProbe;

    protected AbstractDocComConverter(PowerShellRunner powerShellRunner) {
        this.powerShellRunner = powerShellRunner;
    }

    protected abstract String progId();

    protected abstract int[] saveFormatPriority();

    public DocComProbeResult probe() {
        if (!isWindows()) {
            lastProbe = DocComProbeResult.unavailable(engineName(), "仅支持 Windows 环境");
            return lastProbe;
        }
        if (powerShellRunner.getExecutable().isEmpty()) {
            lastProbe = DocComProbeResult.unavailable(engineName(), "未检测到 PowerShell");
            return lastProbe;
        }
        try {
            PowerShellRunner.Result result = powerShellRunner.runScript(buildProbeScript(), PROBE_TIMEOUT);
            lastProbe = new DocComProbeResult(engineName(), true, "可用", result.stdout(), result.stderr(), result.exitCode());
        } catch (PowerShellRunner.PowerShellExecutionException e) {
            PowerShellRunner.Result result = e.getResult();
            lastProbe = new DocComProbeResult(engineName(), false, "COM 探测失败", result.stdout(), result.stderr(), result.exitCode());
        }
        return lastProbe;
    }

    public DocComProbeResult getLastProbe() {
        if (lastProbe == null) {
            return probe();
        }
        return lastProbe;
    }

    @Override
    public boolean isAvailable() {
        return probe().available();
    }

    @Override
    public List<Path> convertBatch(List<Path> docFiles, Path tempDir) throws IOException, DocComConversionException {
        if (docFiles == null || docFiles.isEmpty()) {
            return List.of();
        }
        DocComProbeResult probe = probe();
        if (!probe.available()) {
            throw new DocComConversionException("当前环境无法使用" + engineName() + "进行 .doc 转换：" + probe.message(),
                    docFiles.get(0).toString(), probe.stdout(), probe.stderr(), probe.exitCode());
        }
        if (tempDir == null) {
            throw new IOException("临时目录不能为空");
        }
        Files.createDirectories(tempDir);
        List<Path> outputs = new ArrayList<>();
        try {
            COM_LOCK.acquire();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("转换锁获取失败", e);
        }
        try {
            for (int i = 0; i < docFiles.size(); i++) {
                Path input = docFiles.get(i);
                String outputName = buildOutputName(i, input);
                Path output = tempDir.resolve(outputName);
                runConversion(input, output);
                if (!Files.exists(output)) {
                    throw new DocComConversionException("转换失败，未生成输出文件",
                            input.toString(), "", "未生成输出文件：" + output, -1);
                }
                outputs.add(output);
            }
        } finally {
            COM_LOCK.release();
        }
        return outputs;
    }

    protected void runConversion(Path input, Path output) throws DocComConversionException {
        String script = buildConversionScript(input, output);
        try {
            powerShellRunner.runScript(script, CONVERT_TIMEOUT);
        } catch (PowerShellRunner.PowerShellExecutionException e) {
            PowerShellRunner.Result result = e.getResult();
            throw new DocComConversionException(engineName() + " 转换失败",
                    input.toString(), result.stdout(), result.stderr(), result.exitCode());
        }
    }

    private String buildProbeScript() {
        StringBuilder builder = new StringBuilder();
        builder.append("$ErrorActionPreference = 'Stop'").append(System.lineSeparator());
        builder.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("$app = New-Object -ComObject ").append(progId()).append(System.lineSeparator());
        builder.append("try {").append(System.lineSeparator());
        builder.append("  try { $app.Visible = $false } catch { }").append(System.lineSeparator());
        builder.append("  try { $app.DisplayAlerts = 0 } catch { }").append(System.lineSeparator());
        builder.append("} finally {").append(System.lineSeparator());
        builder.append("  if ($app -ne $null) {").append(System.lineSeparator());
        builder.append("    try { $app.Quit() | Out-Null } catch { }").append(System.lineSeparator());
        builder.append("    [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($app) | Out-Null").append(System.lineSeparator());
        builder.append("  }").append(System.lineSeparator());
        builder.append("}").append(System.lineSeparator());
        return builder.toString();
    }

    private String buildConversionScript(Path input, Path output) {
        int[] priority = saveFormatPriority();
        StringBuilder builder = new StringBuilder();
        builder.append("$ErrorActionPreference = 'Stop'").append(System.lineSeparator());
        builder.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(System.lineSeparator());
        builder.append("$inputPath = '").append(escapePowerShell(input.toString())).append("'").append(System.lineSeparator());
        builder.append("$outputPath = '").append(escapePowerShell(output.toString())).append("'").append(System.lineSeparator());
        builder.append("$app = New-Object -ComObject ").append(progId()).append(System.lineSeparator());
        builder.append("try {").append(System.lineSeparator());
        builder.append("  try { $app.Visible = $false } catch { }").append(System.lineSeparator());
        builder.append("  try { $app.DisplayAlerts = 0 } catch { }").append(System.lineSeparator());
        builder.append("  $doc = $null").append(System.lineSeparator());
        builder.append("  try {").append(System.lineSeparator());
        builder.append("    $doc = $app.Documents.Open($inputPath, $false, $true)").append(System.lineSeparator());
        builder.append("    $saved = $false").append(System.lineSeparator());
        builder.append("    try { $doc.SaveAs2($outputPath, ").append(priority[0]).append("); $saved = $true } catch {").append(System.lineSeparator());
        builder.append("      try { $doc.SaveAs2($outputPath, ").append(priority[1]).append("); $saved = $true } catch {").append(System.lineSeparator());
        builder.append("        try { $doc.SaveAs($outputPath, ").append(priority[0]).append("); $saved = $true } catch {").append(System.lineSeparator());
        builder.append("          $doc.SaveAs($outputPath, ").append(priority[1]).append("); $saved = $true").append(System.lineSeparator());
        builder.append("        }").append(System.lineSeparator());
        builder.append("      }").append(System.lineSeparator());
        builder.append("    }").append(System.lineSeparator());
        builder.append("    if (-not $saved) { throw '保存失败' }").append(System.lineSeparator());
        builder.append("  } finally {").append(System.lineSeparator());
        builder.append("    if ($doc -ne $null) {").append(System.lineSeparator());
        builder.append("      try { $doc.Close([ref]$false) | Out-Null } catch { }").append(System.lineSeparator());
        builder.append("      [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($doc) | Out-Null").append(System.lineSeparator());
        builder.append("    }").append(System.lineSeparator());
        builder.append("  }").append(System.lineSeparator());
        builder.append("} finally {").append(System.lineSeparator());
        builder.append("  if ($app -ne $null) {").append(System.lineSeparator());
        builder.append("    try { $app.Quit() | Out-Null } catch { }").append(System.lineSeparator());
        builder.append("    [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($app) | Out-Null").append(System.lineSeparator());
        builder.append("  }").append(System.lineSeparator());
        builder.append("}").append(System.lineSeparator());
        return builder.toString();
    }

    private String buildOutputName(int index, Path input) {
        String baseName = stripExtension(input.getFileName().toString());
        return String.format(Locale.ROOT, "%03d_%s_%s.docx", index + 1, baseName, UUID.randomUUID());
    }

    private String stripExtension(String name) {
        int idx = name.lastIndexOf('.');
        return idx > 0 ? name.substring(0, idx) : name;
    }

    private String escapePowerShell(String value) {
        return value.replace("'", "''");
    }

    private boolean isWindows() {
        return System.getProperty("os.name").toLowerCase(Locale.ROOT).contains("win");
    }
}
