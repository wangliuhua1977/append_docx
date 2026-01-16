package app.docmerge;

import javax.imageio.ImageIO;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.SwingUtilities;
import javax.swing.filechooser.FileView;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ThumbnailFileView extends FileView {
    private final JFileChooser chooser;
    private final UiLogger logger;
    private final int thumbSize;
    private final Map<String, ImageIcon> cache;
    private final Set<String> pending;
    private final Set<String> warned;
    private final ExecutorService executor;

    public ThumbnailFileView(JFileChooser chooser, UiLogger logger, int thumbSize, int cacheSize) {
        this.chooser = chooser;
        this.logger = logger;
        this.thumbSize = Math.max(16, thumbSize);
        this.cache = Collections.synchronizedMap(new LinkedHashMap<>(16, 0.75f, true) {
            @Override
            protected boolean removeEldestEntry(Map.Entry<String, ImageIcon> eldest) {
                return size() > cacheSize;
            }
        });
        this.pending = ConcurrentHashMap.newKeySet();
        this.warned = ConcurrentHashMap.newKeySet();
        this.executor = Executors.newFixedThreadPool(2, runnable -> {
            Thread thread = new Thread(runnable, "thumbnail-worker");
            thread.setDaemon(true);
            return thread;
        });
    }

    @Override
    public Icon getIcon(File file) {
        if (file == null || file.isDirectory()) {
            return null;
        }
        if (!isImageFile(file)) {
            return null;
        }
        String key = cacheKey(file);
        ImageIcon cached = cache.get(key);
        if (cached != null) {
            return cached;
        }
        if (pending.add(key)) {
            executor.execute(() -> {
                try {
                    ImageIcon icon = loadThumbnail(file);
                    if (icon != null) {
                        cache.put(key, icon);
                    }
                } finally {
                    pending.remove(key);
                    SwingUtilities.invokeLater(chooser::repaint);
                }
            });
        }
        return null;
    }

    private boolean isImageFile(File file) {
        String name = file.getName().toLowerCase(Locale.ROOT);
        return name.endsWith(".png")
                || name.endsWith(".jpg")
                || name.endsWith(".jpeg")
                || name.endsWith(".bmp")
                || name.endsWith(".gif");
    }

    private String cacheKey(File file) {
        return file.getAbsolutePath() + "|" + file.lastModified() + "|" + file.length();
    }

    private ImageIcon loadThumbnail(File file) {
        BufferedImage image;
        try {
            image = ImageIO.read(file);
        } catch (IOException e) {
            logger.warn("读取图片缩略图失败：" + file.getName() + "，原因：" + e.getMessage());
            return null;
        }
        if (image == null) {
            String key = file.getAbsolutePath();
            if (warned.add(key)) {
                logger.warn("无法生成图片缩略图，格式可能不受支持：" + file.getName());
            }
            return null;
        }
        int width = image.getWidth();
        int height = image.getHeight();
        if (width <= 0 || height <= 0) {
            return null;
        }
        double scale = Math.min(1.0, Math.min((double) thumbSize / (double) width, (double) thumbSize / (double) height));
        int targetWidth = Math.max(1, (int) Math.round(width * scale));
        int targetHeight = Math.max(1, (int) Math.round(height * scale));
        Image scaled = image.getScaledInstance(targetWidth, targetHeight, Image.SCALE_SMOOTH);
        BufferedImage canvas = new BufferedImage(thumbSize, thumbSize, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = canvas.createGraphics();
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        int x = (thumbSize - targetWidth) / 2;
        int y = (thumbSize - targetHeight) / 2;
        g2d.drawImage(scaled, x, y, null);
        g2d.dispose();
        return new ImageIcon(canvas);
    }
}
