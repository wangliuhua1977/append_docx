package app.docmerge;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAltChunk;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;

public class MergeService {
    private static final String ALT_CHUNK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    public void merge(List<FileItem> items,
                      Path outputDir,
                      String outputName,
                      DocConverterMode mode,
                      DocComConverterResolver resolver,
                      UiLogger logger,
                      ProgressCallback callback,
                      CancelSignal cancelSignal) throws IOException {
        if (items.isEmpty()) {
            throw new IOException("没有可合并的文件");
        }
        Files.createDirectories(outputDir);
        String fileName = (outputName == null || outputName.isBlank())
                ? defaultOutputName()
                : outputName.trim();
        if (!fileName.toLowerCase(Locale.ROOT).endsWith(".docx")) {
            fileName = fileName + ".docx";
        }
        Path outputFile = outputDir.resolve(fileName);
        Path tempFile = outputDir.resolve(fileName + ".tmp");
        Files.deleteIfExists(tempFile);
        Path tempDir = null;
        Map<Path, Path> convertedMap = new HashMap<>();

        try {
            List<FileItem> docItems = items.stream()
                    .filter(item -> item.getName().toLowerCase(Locale.ROOT).endsWith(".doc"))
                    .toList();
            if (!docItems.isEmpty()) {
                logger.info("DOC 转换模式：" + mode.getLabel());
                DocComConverterResolver.Resolution resolution = resolver.resolve(mode, false);
                DocComConverterSelector.Selection selection = resolution.selection();
                if (selection == null) {
                    logProbeFailure(logger, resolution.probeSummary());
                    throw new IOException(resolution.errorMessage() == null
                            ? "检测到 .doc 文件，但当前模式不可用，已阻止合并。"
                            : resolution.errorMessage());
                }
                logger.info("选择引擎：" + selection.status().engineName());
                tempDir = Files.createTempDirectory("doc-merge-com-");
                if (cancelSignal.isCancelled()) {
                    throw new MergeCancelledException("用户已取消合并");
                }
                List<Path> inputs = docItems.stream().map(FileItem::getPath).toList();
                logger.info("开始批量转换 .doc 文件，共 " + inputs.size() + " 个");
                List<Path> outputs = selection.converter().convertBatch(inputs, tempDir);
                if (outputs.size() != inputs.size()) {
                    throw new DocComConversionException("转换失败，输出文件数量不一致",
                            inputs.get(0).toString(), "", "输出数量=" + outputs.size(), -1);
                }
                for (int i = 0; i < inputs.size(); i++) {
                    convertedMap.put(inputs.get(i), outputs.get(i));
                    logger.info("转换完成：" + inputs.get(i));
                }
                if (cancelSignal.isCancelled()) {
                    throw new MergeCancelledException("用户已取消合并");
                }
            }
            try (XWPFDocument document = new XWPFDocument()) {
                AtomicInteger index = new AtomicInteger();
                for (int i = 0; i < items.size(); i++) {
                    if (cancelSignal.isCancelled()) {
                        throw new MergeCancelledException("用户已取消合并");
                    }
                    FileItem item = items.get(i);
                    callback.onProgress(i + 1, items.size(), item.getName());
                    String lower = item.getName().toLowerCase(Locale.ROOT);
                    if (lower.endsWith(".docx")) {
                        appendDocx(document, item.getPath(), index.incrementAndGet());
                    } else if (lower.endsWith(".doc")) {
                        Path converted = convertedMap.get(item.getPath());
                        if (converted == null) {
                            throw new IOException("未找到 .doc 转换结果：" + item.getPath());
                        }
                        appendDocx(document, converted, index.incrementAndGet());
                    }
                    if (i < items.size() - 1) {
                        XWPFParagraph para = document.createParagraph();
                        para.createRun().addBreak(BreakType.PAGE);
                    }
                }
                try (OutputStream out = Files.newOutputStream(tempFile)) {
                    document.write(out);
                }
            }
            Files.move(tempFile, outputFile, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            logger.info("合并完成，输出文件：" + outputFile);
        } catch (DocComConversionException e) {
            Files.deleteIfExists(tempFile);
            logger.error("COM 转换失败，文件：" + e.getFailedInput()
                    + "，退出码：" + e.getExitCode()
                    + "\nstdout:\n" + e.getStdout()
                    + "\nstderr:\n" + e.getStderr(), e);
            throw new IOException("COM 转换失败：" + e.getMessage(), e);
        } catch (IOException e) {
            Files.deleteIfExists(tempFile);
            throw e;
        } finally {
            cleanupTempDir(tempDir);
        }
    }

    private void appendDocx(XWPFDocument document, Path docxPath, int index) throws IOException {
        try (InputStream in = Files.newInputStream(docxPath)) {
            PackagePartName partName = PackagingURIHelper.createPartName("/word/altChunk" + index + ".docx");
            PackagePart part = document.getPackage().createPart(partName, "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main");
            try (OutputStream partOut = part.getOutputStream()) {
                in.transferTo(partOut);
            }
            String relId = document.getPackagePart().addRelationship(partName, TargetMode.INTERNAL, ALT_CHUNK_REL).getId();
            CTAltChunk chunk = document.getDocument().getBody().addNewAltChunk();
            chunk.setId(relId);
        } catch (InvalidFormatException e) {
            throw new IOException("插入 altChunk 失败：" + docxPath.getFileName(), e);
        }
    }

    private String defaultOutputName() {
        return "合并结果_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss")) + ".docx";
    }

    @FunctionalInterface
    public interface ProgressCallback {
        void onProgress(int current, int total, String name);
    }

    public interface CancelSignal {
        boolean isCancelled();
    }

    public static class MergeCancelledException extends IOException {
        public MergeCancelledException(String message) {
            super(message);
        }
    }

    private void cleanupTempDir(Path tempDir) {
        if (tempDir == null) {
            return;
        }
        try {
            Files.walk(tempDir)
                    .sorted((a, b) -> b.getNameCount() - a.getNameCount())
                    .forEach(path -> {
                        try {
                            Files.deleteIfExists(path);
                        } catch (IOException ignored) {
                            // ignore
                        }
                    });
        } catch (IOException ignored) {
            // ignore
        }
    }

    private void logProbeFailure(UiLogger logger, DocComConverterSelector.ProbeSummary probeSummary) {
        if (probeSummary == null) {
            logger.warn("COM 探测失败：未知原因");
            return;
        }
        DocComConverterSelector.EngineStatus word = probeSummary.word();
        DocComConverterSelector.EngineStatus wps = probeSummary.wps();
        if (word != null) {
            logger.warn("Word 探测不可用：" + word.message()
                    + "，exitCode=" + word.exitCode()
                    + "\nstderr:\n" + word.stderr());
        }
        if (wps != null) {
            logger.warn("WPS 探测不可用：" + wps.message()
                    + "，exitCode=" + wps.exitCode()
                    + "\nstderr:\n" + wps.stderr());
        }
    }
}
