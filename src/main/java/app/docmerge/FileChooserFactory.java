package app.docmerge;

import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;

public final class FileChooserFactory {
    private static final int THUMBNAIL_SIZE = 64;
    private static final int THUMBNAIL_CACHE_SIZE = 256;

    private FileChooserFactory() {
    }

    public static JFileChooser createAddFilesChooser(UiLogger logger) {
        JFileChooser chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        chooser.setMultiSelectionEnabled(true);
        chooser.setAcceptAllFileFilterUsed(true);
        chooser.setFileFilter(new FileNameExtensionFilter(
                "支持的文件 (*.doc, *.docx, *.pdf, *.png, *.jpg, *.jpeg, *.bmp, *.gif)",
                "doc", "docx", "pdf", "png", "jpg", "jpeg", "bmp", "gif"));
        chooser.setFileView(new ThumbnailFileView(chooser, logger, THUMBNAIL_SIZE, THUMBNAIL_CACHE_SIZE));
        return chooser;
    }
}
