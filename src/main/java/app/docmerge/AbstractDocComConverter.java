// AbstractDocComConverter.java
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
    private static final Duration PDF_CONVERT_TIMEOUT = Duration.ofMinutes(15);

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

    @Override
    public boolean supportsPdfConversion() {
        return true;
    }

    @Override
    public Path convertPdfToDocx(Path pdfFile, Path tempDir) throws IOException, DocComConversionException {
        if (pdfFile == null) {
            throw new IOException("PDF 文件为空");
        }
        DocComProbeResult probe = probe();
        if (!probe.available()) {
            throw new DocComConversionException("当前环境无法使用" + engineName() + "进行 PDF 转换：" + probe.message(),
                    pdfFile.toString(), probe.stdout(), probe.stderr(), probe.exitCode());
        }
        if (!supportsPdfConversion()) {
            throw new DocComConversionException("当前引擎不支持 PDF 转 DOCX",
                    pdfFile.toString(), "", "当前引擎不支持 PDF 转 DOCX", -1);
        }
        if (tempDir == null) {
            throw new IOException("临时目录不能为空");
        }
        Files.createDirectories(tempDir);

        try {
            COM_LOCK.acquire();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            throw new IOException("转换锁获取失败", e);
        }

        try {
            String outputName = buildOutputName(0, pdfFile);
            Path output = tempDir.resolve(outputName);
            runPdfConversion(pdfFile, output);

            if (!Files.exists(output)) {
                throw new DocComConversionException("转换失败，未生成输出文件",
                        pdfFile.toString(), "", "未生成输出文件：" + output, -1);
            }
            return output;
        } finally {
            COM_LOCK.release();
        }
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

    protected void runPdfConversion(Path input, Path output) throws DocComConversionException {
        String script = buildPdfConversionScript(input, output);
        try {
            powerShellRunner.runScript(script, PDF_CONVERT_TIMEOUT);
        } catch (PowerShellRunner.PowerShellExecutionException e) {
            PowerShellRunner.Result result = e.getResult();
            throw new DocComConversionException(engineName() + " PDF 转换失败",
                    input.toString(), result.stdout(), result.stderr(), result.exitCode());
        }
    }

    private String buildProbeScript() {
        String ls = System.lineSeparator();
        StringBuilder b = new StringBuilder();
        b.append("$ErrorActionPreference = 'Stop'").append(ls);
        b.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("function _err([string]$m) { try { [Console]::Error.WriteLine($m) } catch { } }").append(ls);

        b.append("try {").append(ls);
        b.append("  $app = $null").append(ls);
        b.append("  $app = New-Object -ComObject ").append(progId()).append(ls);
        b.append("  try {").append(ls);
        b.append("    try { $app.Visible = $false } catch { }").append(ls);
        b.append("    try { $app.DisplayAlerts = 0 } catch { }").append(ls);
        b.append("  } finally {").append(ls);
        b.append("    if ($app -ne $null) {").append(ls);
        b.append("      try { $app.Quit() | Out-Null } catch { }").append(ls);
        b.append("      try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("  }").append(ls);
        b.append("  exit 0").append(ls);
        b.append("} catch {").append(ls);
        // 关键：禁止 Write-Error（Stop 下会二次抛异常），用 Console.Error 输出
        b.append("  _err ('[COM探测异常] ' + $_.Exception.ToString())").append(ls);
        b.append("  try { _err ('[ErrorRecord] ' + ($_ | Format-List -Force * | Out-String)) } catch { }").append(ls);
        b.append("  exit 1").append(ls);
        b.append("}").append(ls);
        return b.toString();
    }

    private String buildConversionScript(Path input, Path output) {
        int[] priority = saveFormatPriority();
        String ls = System.lineSeparator();

        StringBuilder b = new StringBuilder();
        b.append("$ErrorActionPreference = 'Stop'").append(ls);
        b.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("function _err([string]$m) { try { [Console]::Error.WriteLine($m) } catch { } }").append(ls);

        b.append("$inputPath = '").append(escapePowerShell(input.toString())).append("'").append(ls);
        b.append("$outputPath = '").append(escapePowerShell(output.toString())).append("'").append(ls);

        b.append("try {").append(ls);
        b.append("  if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }").append(ls);

        b.append("  $app = $null").append(ls);
        b.append("  $doc = $null").append(ls);
        b.append("  $pv  = $null").append(ls);

        b.append("  $app = New-Object -ComObject ").append(progId()).append(ls);
        b.append("  try {").append(ls);
        b.append("    try { $app.Visible = $false } catch { }").append(ls);
        b.append("    try { $app.DisplayAlerts = 0 } catch { }").append(ls);
        b.append("    try { $app.AutomationSecurity = 3 } catch { }").append(ls);
        b.append("    try { $app.Options.ConfirmConversions = $false } catch { }").append(ls);

        // Open：参数签名回退 + ProtectedView 回退（兼容 Word/WPS）
        b.append("    try {").append(ls);
        b.append("      $doc = $app.Documents.Open($inputPath, $false, $true, $false)").append(ls);
        b.append("    } catch {").append(ls);
        b.append("      try {").append(ls);
        b.append("        $pv = $app.ProtectedViewWindows.Open($inputPath)").append(ls);
        b.append("        $doc = $pv.Edit()").append(ls);
        b.append("      } catch {").append(ls);
        b.append("        $doc = $app.Documents.Open($inputPath)").append(ls);
        b.append("      }").append(ls);
        b.append("    }").append(ls);

        b.append("    if ($doc -eq $null) { throw '无法打开文档：' + $inputPath }").append(ls);

        // Save：SaveAs2/SaveAs + 多格式优先级 + 最后不带格式
        b.append("    $saved = $false").append(ls);
        b.append("    $fmts = @(").append(priority[0]).append(", ").append(priority[1]).append(")").append(ls);
        b.append("    foreach ($fmt in $fmts) {").append(ls);
        b.append("      if ($saved) { break }").append(ls);
        b.append("      try { $doc.SaveAs2($outputPath, [int]$fmt); $saved = $true } catch { }").append(ls);
        b.append("      if (-not $saved) { try { $doc.SaveAs($outputPath, [int]$fmt); $saved = $true } catch { } }").append(ls);
        b.append("    }").append(ls);
        b.append("    if (-not $saved) {").append(ls);
        b.append("      try { $doc.SaveAs($outputPath); $saved = $true } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("    if (-not $saved) { throw '保存失败' }").append(ls);

        b.append("  } finally {").append(ls);
        b.append("    if ($doc -ne $null) {").append(ls);
        b.append("      try { $doc.Close($false) | Out-Null } catch { }").append(ls);
        b.append("      try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($doc) | Out-Null } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("    if ($pv -ne $null) { try { $pv.Close() | Out-Null } catch { } }").append(ls);
        b.append("    if ($app -ne $null) {").append(ls);
        b.append("      try { $app.Quit() | Out-Null } catch { }").append(ls);
        b.append("      try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("    [GC]::Collect()").append(ls);
        b.append("    [GC]::WaitForPendingFinalizers()").append(ls);
        b.append("  }").append(ls);

        b.append("  exit 0").append(ls);
        b.append("} catch {").append(ls);
        b.append("  _err ('[DOC转DOCX异常] ' + $_.Exception.ToString())").append(ls);
        b.append("  try { _err ('[ErrorRecord] ' + ($_ | Format-List -Force * | Out-String)) } catch { }").append(ls);
        b.append("  exit 1").append(ls);
        b.append("}").append(ls);

        return b.toString();
    }

    /**
     * PDF 转 DOCX（Word）增强版：
     * - 关键修复：catch 内不用 Write-Error（Stop 下会二次终止 -> exitCode 常见变成 -1），改 Console.Error 输出
     * - Open 回退：Documents.Open(参数) -> ProtectedView -> Documents.Open(单参)
     * - SaveAs2 重试 + 多格式（优先 16，再 12）+ SaveAs 回退
     * - PDF 场景下临时 Visible=true（部分环境 PDF Reflow 需要可见 UI 才稳定）
     */
    private String buildPdfConversionScript(Path input, Path output) {
        int[] priority = saveFormatPriority();
        String ls = System.lineSeparator();
        StringBuilder b = new StringBuilder();

        b.append("$ErrorActionPreference = 'Stop'").append(ls);
        b.append("[Console]::OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("$OutputEncoding = [System.Text.Encoding]::UTF8").append(ls);
        b.append("function _err([string]$m) { try { [Console]::Error.WriteLine($m) } catch { } }").append(ls);

        b.append("$inputPath = '").append(escapePowerShell(input.toString())).append("'").append(ls);
        b.append("$outputPath = '").append(escapePowerShell(output.toString())).append("'").append(ls);

        b.append("try {").append(ls);
        b.append("  if (Test-Path -LiteralPath $outputPath) { Remove-Item -LiteralPath $outputPath -Force }").append(ls);

        b.append("  $app = $null").append(ls);
        b.append("  $doc = $null").append(ls);
        b.append("  $pv  = $null").append(ls);

        b.append("  $app = New-Object -ComObject ").append(progId()).append(ls);
        b.append("  try {").append(ls);
        // PDF Reflow 在部分环境需要可见窗口更稳定
        b.append("    try { $app.Visible = $true } catch { }").append(ls);
        b.append("    try { $app.DisplayAlerts = 0 } catch { }").append(ls);
        b.append("    try { $app.AutomationSecurity = 3 } catch { }").append(ls);
        b.append("    try { $app.Options.ConfirmConversions = $false } catch { }").append(ls);

        // Open 回退链
        b.append("    try {").append(ls);
        b.append("      $doc = $app.Documents.Open($inputPath, $false, $true, $false)").append(ls);
        b.append("    } catch {").append(ls);
        b.append("      $doc = $null").append(ls);
        b.append("    }").append(ls);

        b.append("    if ($doc -eq $null) {").append(ls);
        b.append("      try {").append(ls);
        b.append("        $pv = $app.ProtectedViewWindows.Open($inputPath)").append(ls);
        b.append("        $doc = $pv.Edit()").append(ls);
        b.append("      } catch {").append(ls);
        b.append("        $doc = $null").append(ls);
        b.append("      }").append(ls);
        b.append("    }").append(ls);

        b.append("    if ($doc -eq $null) {").append(ls);
        b.append("      $doc = $app.Documents.Open($inputPath)").append(ls);
        b.append("    }").append(ls);

        // 兜底：某些版本 Edit() 返回 null，但 ActiveDocument 已经有值
        b.append("    if ($doc -eq $null) { try { $doc = $app.ActiveDocument } catch { } }").append(ls);
        b.append("    if ($doc -eq $null) { throw '无法打开 PDF：' + $inputPath }").append(ls);

        // 给 PDF Reflow 一点缓冲，并尝试重排版
        b.append("    try { $doc.Activate() | Out-Null } catch { }").append(ls);
        b.append("    Start-Sleep -Milliseconds 800").append(ls);
        b.append("    try { $doc.Repaginate() | Out-Null } catch { }").append(ls);

        // Save 多格式 + 重试
        b.append("    $saved = $false").append(ls);

        // PDF 场景：优先 16，再 12（兼容部分 Word/WPS/兼容模式行为差异）
        b.append("    $fmts = @(").append(priority[0]).append(", ").append(priority[1]).append(")").append(ls);

        b.append("    foreach ($fmt in $fmts) {").append(ls);
        b.append("      if ($saved) { break }").append(ls);

        // SaveAs2 重试三次（2/4/6 秒）
        b.append("      for ($i = 1; $i -le 3; $i++) {").append(ls);
        b.append("        try {").append(ls);
        b.append("          $doc.SaveAs2($outputPath, [int]$fmt)").append(ls);
        b.append("          $saved = $true").append(ls);
        b.append("          break").append(ls);
        b.append("        } catch {").append(ls);
        b.append("          Start-Sleep -Seconds (2 * $i)").append(ls);
        b.append("        }").append(ls);
        b.append("      }").append(ls);

        b.append("      if (-not $saved) {").append(ls);
        b.append("        try { $doc.SaveAs($outputPath, [int]$fmt); $saved = $true } catch { }").append(ls);
        b.append("      }").append(ls);
        b.append("    }").append(ls);

        b.append("    if (-not $saved) {").append(ls);
        b.append("      try { $doc.SaveAs($outputPath); $saved = $true } catch { }").append(ls);
        b.append("    }").append(ls);

        b.append("    if (-not $saved) { throw '保存失败：PDF 转 DOCX 未成功执行 SaveAs/SaveAs2' }").append(ls);

        b.append("  } finally {").append(ls);
        b.append("    if ($doc -ne $null) {").append(ls);
        b.append("      try { $doc.Close($false) | Out-Null } catch { }").append(ls);
        b.append("      try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($doc) | Out-Null } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("    if ($pv -ne $null) { try { $pv.Close() | Out-Null } catch { } }").append(ls);
        b.append("    if ($app -ne $null) {").append(ls);
        b.append("      try { $app.Quit() | Out-Null } catch { }").append(ls);
        b.append("      try { [System.Runtime.Interopservices.Marshal]::FinalReleaseComObject($app) | Out-Null } catch { }").append(ls);
        b.append("    }").append(ls);
        b.append("    [GC]::Collect()").append(ls);
        b.append("    [GC]::WaitForPendingFinalizers()").append(ls);
        b.append("  }").append(ls);

        b.append("  exit 0").append(ls);
        b.append("} catch {").append(ls);
        b.append("  _err ('[PDF转DOCX异常] ' + $_.Exception.ToString())").append(ls);
        b.append("  try { _err ('[ErrorRecord] ' + ($_ | Format-List -Force * | Out-String)) } catch { }").append(ls);
        // 明确 exit 1，避免出现 0xFFFFFFFF -> Java 显示 -1
        b.append("  exit 1").append(ls);
        b.append("}").append(ls);

        return b.toString();
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
