package app.docmerge;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
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
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;

public class MergeService {
    private static final String ALT_CHUNK_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk";
    private final LibreOfficeConverter libreOfficeConverter = new LibreOfficeConverter();

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
                    boolean handled = convertAndAppendDoc(document, item, index.incrementAndGet(), logger);
                    if (!handled) {
                        appendDocFallback(document, item, logger);
                    }
                }
                if (i < items.size() - 1) {
                    XWPFParagraph para = document.createParagraph();
                    para.createRun().addBreak(BreakType.PAGE);
                }
            }
            try (OutputStream out = Files.newOutputStream(tempFile)) {
                document.write(out);
            }
            Files.move(tempFile, outputFile, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            logger.info("合并完成，输出文件：" + outputFile);
        } catch (IOException e) {
            Files.deleteIfExists(tempFile);
            throw e;
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

    private boolean convertAndAppendDoc(XWPFDocument document, FileItem item, int index, UiLogger logger) {
        Path tempDir = Path.of(System.getProperty("java.io.tmpdir"), "doc-merge-app", "lo" + System.nanoTime());
        Optional<Path> converted = libreOfficeConverter.convertToDocx(item.getPath(), tempDir, logger);
        if (converted.isPresent()) {
            try {
                appendDocx(document, converted.get(), index);
                return true;
            } catch (IOException e) {
                logger.error("插入 LibreOffice 转换结果失败：" + item.getName(), e);
                return false;
            } finally {
                try {
                    Files.deleteIfExists(converted.get());
                } catch (IOException ignored) {
                    // ignore
                }
            }
        }
        return false;
    }

    private void appendDocFallback(XWPFDocument document, FileItem item, UiLogger logger) {
        logger.warn("未检测到 LibreOffice 或转换失败，使用 .doc 兼容模式：" + item.getName() + "，格式可能无法完全保留");
        try (InputStream in = Files.newInputStream(item.getPath());
             HWPFDocument hwpf = new HWPFDocument(in)) {
            WordExtractor extractor = new WordExtractor(hwpf);
            String[] paragraphs = extractor.getParagraphText();
            for (String text : paragraphs) {
                XWPFParagraph para = document.createParagraph();
                para.createRun().setText(text.replace("\r", "").replace("\n", ""));
            }
        } catch (IOException e) {
            logger.error("读取 .doc 失败：" + item.getName(), e);
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
}
