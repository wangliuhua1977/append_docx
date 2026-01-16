package app.docmerge;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackagePartName;
import org.apache.poi.openxml4j.opc.PackagingURIHelper;
import org.apache.poi.openxml4j.opc.TargetMode;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTAltChunk;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
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
    private static final int IMAGE_DPI = 96;
    private static final long TWIP_TO_EMU = 635L;
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
        Map<Path, Path> convertedPdfMap = new HashMap<>();
        int totalUnits = Math.max(items.size(), 1);
        long startTime = System.currentTimeMillis();
        DocComConverterSelector.Selection selection = null;

        try {
            List<FileItem> docItems = items.stream()
                    .filter(item -> item.getFileType() == FileItem.FileType.DOC)
                    .toList();
            List<FileItem> pdfItems = items.stream()
                    .filter(item -> item.getFileType() == FileItem.FileType.PDF)
                    .toList();
            if (!docItems.isEmpty() || !pdfItems.isEmpty()) {
                logger.info("DOC/PDF 转换模式：" + mode.getLabel());
                DocComConverterResolver.Resolution resolution = resolver.resolve(mode, false);
                selection = resolution.selection();
                if (selection == null) {
                    logProbeFailure(logger, resolution.probeSummary());
                    throw new IOException(resolution.errorMessage() == null
                            ? "检测到 .doc 或 PDF 文件，但当前模式不可用，已阻止合并。"
                            : resolution.errorMessage());
                }
                if (!pdfItems.isEmpty() && !selection.converter().supportsPdfConversion()) {
                    throw new IOException("当前引擎不支持 PDF 转 DOCX：" + selection.status().engineName());
                }
                logger.info("选择引擎：" + selection.status().engineName());
                tempDir = Files.createTempDirectory("doc-merge-com-");
                if (cancelSignal.isCancelled()) {
                    throw new MergeCancelledException("用户已取消合并");
                }
                if (!docItems.isEmpty()) {
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
                }
                if (cancelSignal.isCancelled()) {
                    throw new MergeCancelledException("用户已取消合并");
                }
            }
            try (XWPFDocument document = new XWPFDocument()) {
                AtomicInteger index = new AtomicInteger();
                ProgressTracker tracker = new ProgressTracker(totalUnits, callback);
                for (int i = 0; i < items.size(); i++) {
                    if (cancelSignal.isCancelled()) {
                        throw new MergeCancelledException("用户已取消合并");
                    }
                    FileItem item = items.get(i);
                    logger.info("开始处理：" + item.getName() + "，类型：" + item.getFileType().getLabel());
                    switch (item.getFileType()) {
                        case DOCX -> {
                            appendDocx(document, item.getPath(), index.incrementAndGet());
                            tracker.step(item.getName());
                            if (i < items.size() - 1) {
                                addPageBreak(document);
                            }
                        }
                        case DOC -> {
                            Path converted = convertedMap.get(item.getPath());
                            if (converted == null) {
                                throw new IOException("未找到 .doc 转换结果：" + item.getPath());
                            }
                            appendDocx(document, converted, index.incrementAndGet());
                            tracker.step(item.getName());
                            if (i < items.size() - 1) {
                                addPageBreak(document);
                            }
                        }
                        case IMAGE -> {
                            appendImageToDocx(document, item.getPath(), logger, tracker::step);
                        }
                        case PDF -> {
                            if (selection == null) {
                                throw new IOException("未选择 PDF 转换引擎：" + item.getPath());
                            }
                            Path converted = convertedPdfMap.get(item.getPath());
                            if (converted == null) {
                                logger.info("开始转换 PDF：" + item.getName());
                                converted = selection.converter().convertPdfToDocx(item.getPath(), tempDir);
                                convertedPdfMap.put(item.getPath(), converted);
                                logger.info("PDF 转换完成：" + item.getName());
                            }
                            appendDocx(document, converted, index.incrementAndGet());
                            tracker.step(item.getName());
                            if (i < items.size() - 1) {
                                addPageBreak(document);
                            }
                        }
                    }
                }
                try (OutputStream out = Files.newOutputStream(tempFile)) {
                    document.write(out);
                }
            }
            Files.move(tempFile, outputFile, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            long cost = System.currentTimeMillis() - startTime;
            logger.info("合并完成，耗时 " + cost + " ms，输出文件：" + outputFile);
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

    private void appendImageToDocx(XWPFDocument document,
                                   Path imagePath,
                                   UiLogger logger,
                                   ProgressStepCallback stepCallback) throws IOException {
        BufferedImage image;
        try {
            image = ImageIO.read(imagePath.toFile());
        } catch (IOException e) {
            logger.error("读取图片失败：" + imagePath, e);
            throw e;
        }
        if (image == null) {
            IOException e = new IOException("无法识别的图片格式：" + imagePath);
            logger.error("读取图片失败：" + imagePath, e);
            throw e;
        }

        long widthEmu = toEmuFromPixels(image.getWidth(), IMAGE_DPI);
        long heightEmu = toEmuFromPixels(image.getHeight(), IMAGE_DPI);
        long maxWidthEmu = resolveUsablePageWidthEmu(document);
        if (widthEmu > maxWidthEmu) {
            double scale = (double) maxWidthEmu / (double) widthEmu;
            widthEmu = Math.round(widthEmu * scale);
            heightEmu = Math.round(heightEmu * scale);
        }

        XWPFParagraph imagePara = document.createParagraph();
        imagePara.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = imagePara.createRun();
        try (InputStream in = Files.newInputStream(imagePath)) {
            run.addPicture(in, pictureType(imagePath), "",
                    safeEmu(widthEmu), safeEmu(heightEmu));
        } catch (Exception e) {
            logger.error("插入图片失败：" + imagePath, e);
            throw new IOException("插入图片失败：" + imagePath.getFileName(), e);
        }
        addPageBreak(document);
        stepCallback.step(imagePath.getFileName().toString());
        logger.info("图片插入完成：" + imagePath.getFileName());
    }

    private void addPageBreak(XWPFDocument document) {
        XWPFParagraph para = document.createParagraph();
        para.createRun().addBreak(BreakType.PAGE);
    }

    private int pictureType(Path imagePath) throws IOException {
        String lower = imagePath.getFileName().toString().toLowerCase(Locale.ROOT);
        if (lower.endsWith(".png")) {
            return XWPFDocument.PICTURE_TYPE_PNG;
        }
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
            return XWPFDocument.PICTURE_TYPE_JPEG;
        }
        if (lower.endsWith(".bmp")) {
            return XWPFDocument.PICTURE_TYPE_BMP;
        }
        if (lower.endsWith(".gif")) {
            return XWPFDocument.PICTURE_TYPE_GIF;
        }
        throw new IOException("不支持的图片格式：" + imagePath.getFileName());
    }

    private long resolveUsablePageWidthEmu(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        long pageWidthTwips = 12240L;
        long marginLeftTwips = 1440L;
        long marginRightTwips = 1440L;
        if (sectPr != null) {
            if (sectPr.isSetPgSz()) {
                CTPageSz pageSz = sectPr.getPgSz();
                Object widthValue = pageSz.getW();
                if (widthValue instanceof Number number) {
                    pageWidthTwips = number.longValue();
                }
            }
            if (sectPr.isSetPgMar()) {
                CTPageMar pageMar = sectPr.getPgMar();
                Object leftValue = pageMar.getLeft();
                if (leftValue instanceof Number number) {
                    marginLeftTwips = number.longValue();
                }
                Object rightValue = pageMar.getRight();
                if (rightValue instanceof Number number) {
                    marginRightTwips = number.longValue();
                }
            }
        }
        long widthTwips = Math.max(1, pageWidthTwips - marginLeftTwips - marginRightTwips);
        return widthTwips * TWIP_TO_EMU;
    }

    private long toEmuFromPixels(int pixels, int dpi) {
        double inches = (double) pixels / (double) dpi;
        double points = inches * 72.0;
        return Math.round(Units.toEMU(points));
    }

    private int safeEmu(long emu) {
        if (emu > Integer.MAX_VALUE) {
            return Integer.MAX_VALUE;
        }
        if (emu < 0) {
            return 0;
        }
        return (int) emu;
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

    @FunctionalInterface
    private interface ProgressStepCallback {
        void step(String name);
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

    private static class ProgressTracker {
        private final int total;
        private final ProgressCallback callback;
        private int current;

        private ProgressTracker(int total, ProgressCallback callback) {
            this.total = Math.max(total, 1);
            this.callback = callback;
            this.current = 0;
        }

        private void step(String name) {
            current += 1;
            callback.onProgress(current, total, name);
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
