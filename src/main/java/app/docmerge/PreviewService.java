package app.docmerge;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.imageio.ImageIO;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.text.DecimalFormat;
import java.util.LinkedHashMap;
import java.util.Map;

public class PreviewService {
    public static final int MAX_TEXT_LENGTH = 20_000;
    public static final int MAX_IMAGE_DIMENSION = 480;
    private static final int MAX_PREVIEW_CACHE = 50;
    private static final int MAX_CONVERT_CACHE = 50;
    private static final DecimalFormat SIZE_FORMAT = new DecimalFormat("#,##0.##");

    private final Map<PreviewKey, PreviewResult> previewCache = new LruCache<>(MAX_PREVIEW_CACHE, null);
    private final Map<PreviewKey, Path> docxCache = new LruCache<>(MAX_CONVERT_CACHE, this::deleteQuietly);
    private final Map<PreviewKey, Path> pdfDocxCache = new LruCache<>(MAX_CONVERT_CACHE, this::deleteQuietly);

    public PreviewResult loadPreview(FileItem item,
                                     DocConverterMode mode,
                                     DocComConverterResolver resolver,
                                     UiLogger logger) {
        if (item == null || item.getPath() == null) {
            return PreviewResult.error("请选择文件以预览");
        }
        PreviewKey key = PreviewKey.fromItem(item);
        PreviewResult cached = previewCache.get(key);
        if (cached != null) {
            return cached;
        }
        try {
            PreviewResult result = switch (item.getFileType()) {
                case IMAGE -> loadImagePreview(item, key);
                case DOCX -> loadDocxPreview(item, key);
                case DOC -> loadDocPreview(item, key, mode, resolver, logger);
                case PDF -> loadPdfPreview(item, key);
            };
            previewCache.put(key, result);
            return result;
        } catch (Exception e) {
            return PreviewResult.error("预览失败：" + e.getMessage());
        }
    }

    private PreviewResult loadImagePreview(FileItem item, PreviewKey key) throws IOException {
        BufferedImage image = ImageIO.read(item.getPath().toFile());
        if (image == null) {
            return PreviewResult.error("无法识别图片格式：" + item.getName());
        }
        BufferedImage scaled = scaleImage(image, MAX_IMAGE_DIMENSION);
        String info = "文件名：" + item.getName()
                + "\n分辨率：" + image.getWidth() + " x " + image.getHeight()
                + "\n大小：" + formatSize(item.getSize());
        return PreviewResult.image(scaled, info);
    }

    private PreviewResult loadDocxPreview(FileItem item, PreviewKey key) throws IOException {
        String text = extractDocxText(item.getPath());
        return PreviewResult.text(text, "文件名：" + item.getName() + "\n类型：DOCX");
    }

    private PreviewResult loadDocPreview(FileItem item,
                                         PreviewKey key,
                                         DocConverterMode mode,
                                         DocComConverterResolver resolver,
                                         UiLogger logger) throws IOException {
        Path converted = docxCache.get(key);
        if (converted == null || !Files.exists(converted)) {
            DocComConverterSelector.Selection selection = resolver.resolve(mode, false).selection();
            if (selection == null) {
                return PreviewResult.error("当前环境无法预览 .doc，请选择 Word/WPS 转换引擎");
            }
            Path tempDir = Files.createTempDirectory("doc-preview-");
            var outputs = selection.converter().convertBatch(java.util.List.of(item.getPath()), tempDir);
            if (outputs.isEmpty()) {
                return PreviewResult.error("DOC 转换失败，无法预览：" + item.getName());
            }
            converted = outputs.get(0);
            docxCache.put(key, converted);
            logger.info("预览转换完成：" + item.getName());
        }
        String text = extractDocxText(converted);
        return PreviewResult.text(text, "文件名：" + item.getName() + "\n类型：DOC（已转为 DOCX 预览）");
    }

    private PreviewResult loadPdfPreview(FileItem item, PreviewKey key) throws IOException {
        try (PDDocument doc = Loader.loadPDF(item.getPath().toFile())) {
            int pages = doc.getNumberOfPages();
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(doc);
            text = truncateText(text);
            String meta = "文件名：" + item.getName() + "\n类型：PDF\n页数：" + pages;
            return PreviewResult.text(text, meta);
        }
    }

    public Path cachePdfConversion(FileItem item, Path converted) {
        if (item == null || converted == null) {
            return converted;
        }
        PreviewKey key = PreviewKey.fromItem(item);
        pdfDocxCache.put(key, converted);
        return converted;
    }

    public Path getCachedPdfConversion(FileItem item) {
        if (item == null) {
            return null;
        }
        return pdfDocxCache.get(PreviewKey.fromItem(item));
    }

    private String extractDocxText(Path path) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(Files.newInputStream(path));
             XWPFWordExtractor extractor = new XWPFWordExtractor(doc)) {
            String text = extractor.getText();
            return truncateText(text);
        }
    }

    private String truncateText(String text) {
        if (text == null) {
            return "";
        }
        if (text.length() <= MAX_TEXT_LENGTH) {
            return text;
        }
        return text.substring(0, MAX_TEXT_LENGTH) + "\n\n【预览已截断，最多 " + MAX_TEXT_LENGTH + " 字】";
    }

    private BufferedImage scaleImage(BufferedImage source, int maxDimension) {
        int width = source.getWidth();
        int height = source.getHeight();
        int max = Math.max(width, height);
        if (max <= maxDimension) {
            return source;
        }
        double scale = (double) maxDimension / (double) max;
        int targetW = Math.max(1, (int) Math.round(width * scale));
        int targetH = Math.max(1, (int) Math.round(height * scale));
        Image scaled = source.getScaledInstance(targetW, targetH, Image.SCALE_SMOOTH);
        BufferedImage result = new BufferedImage(targetW, targetH, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = result.createGraphics();
        g2.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2.drawImage(scaled, 0, 0, null);
        g2.dispose();
        return result;
    }

    private String formatSize(long size) {
        if (size <= 0) {
            return "-";
        }
        double value = size;
        String unit = "B";
        if (value >= 1024) {
            value /= 1024.0;
            unit = "KB";
        }
        if (value >= 1024) {
            value /= 1024.0;
            unit = "MB";
        }
        return SIZE_FORMAT.format(value) + " " + unit;
    }

    private void deleteQuietly(Path path) {
        if (path == null) {
            return;
        }
        try {
            Files.deleteIfExists(path);
        } catch (IOException ignored) {
            // ignore
        }
    }

    public record PreviewResult(PreviewType type, BufferedImage image, String text, String meta) {
        public static PreviewResult image(BufferedImage image, String meta) {
            return new PreviewResult(PreviewType.IMAGE, image, "", meta);
        }

        public static PreviewResult text(String text, String meta) {
            return new PreviewResult(PreviewType.TEXT, null, text, meta);
        }

        public static PreviewResult error(String message) {
            return new PreviewResult(PreviewType.TEXT, null, message, "预览错误");
        }
    }

    public enum PreviewType {
        IMAGE,
        TEXT
    }

    private record PreviewKey(Path path, long lastModified, long size) {
        public static PreviewKey fromItem(FileItem item) {
            FileTime time = item.getLastModified();
            long modified = time == null ? 0L : time.toMillis();
            long size = item.getSize();
            return new PreviewKey(item.getPath().toAbsolutePath(), modified, size);
        }
    }

    private static class LruCache<K, V> extends LinkedHashMap<K, V> {
        private final int maxSize;
        private final java.util.function.Consumer<V> onEvict;

        private LruCache(int maxSize, java.util.function.Consumer<V> onEvict) {
            super(16, 0.75f, true);
            this.maxSize = maxSize;
            this.onEvict = onEvict;
        }

        @Override
        protected boolean removeEldestEntry(Map.Entry<K, V> eldest) {
            boolean remove = size() > maxSize;
            if (remove && onEvict != null) {
                onEvict.accept(eldest.getValue());
            }
            return remove;
        }
    }
}
