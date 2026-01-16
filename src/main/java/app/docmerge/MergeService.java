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
    private final DocComConverterSelector converterSelector = new DocComConverterSelector();

    public void merge(List<FileItem> items,
                      Path outputDir,
                      String outputName,
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
                DocComConverterSelector.ProbeSummary probeSummary = converterSelector.probeAll();
                DocComConverterSelector.Selection selection = probeSummary.selection();
                if (selection == null) {
                    logProbeFailure(logger, probeSummary);
                    throw new IOException("检测到 .doc 文件，但 Word/WPS COM 不可用，已阻止合并。");
                }
                logger.info("已选择转换引擎：" + selection.status().engineName());
                tempDir = Files.createTempDirectory("doc-merge-com-");
                for (int i = 0; i < docItems.size(); i++) {
                    if (cancelSignal.isCancelled()) {
                        throw new MergeCancelledException("用户已取消合并");
                    }
                    FileItem item = docItems.get(i);
                    logger.info("开始转换：" + item.getPath());
                    List<Path> outputs = selection.converter().convertBatch(List.of(item.getPath()), tempDir);
                    if (outputs.isEmpty()) {
                        throw new DocComConversionException("转换失败，未生成输出文件",
                                item.getPath().toString(), "", "未生成输出文件", -1);
                    }
                    Path output = outputs.get(0);
                    convertedMap.put(item.getPath(), output);
                    logger.info("转换完成：" + item.getPath());
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
