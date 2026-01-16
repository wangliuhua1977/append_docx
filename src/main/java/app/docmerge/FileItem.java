package app.docmerge;

import java.nio.file.Path;
import java.nio.file.attribute.FileTime;

public class FileItem {
    public enum FileType {
        DOC("DOC"),
        DOCX("DOCX"),
        IMAGE("图片"),
        PDF("PDF");

        private final String label;

        FileType(String label) {
            this.label = label;
        }

        public String getLabel() {
            return label;
        }

        public static FileType fromExtension(String extension) {
            if (extension == null) {
                return DOCX;
            }
            String lower = extension.toLowerCase();
            return switch (lower) {
                case "doc" -> DOC;
                case "docx" -> DOCX;
                case "pdf" -> PDF;
                case "png", "jpg", "jpeg", "bmp", "gif" -> IMAGE;
                default -> DOCX;
            };
        }
    }

    public enum Status {
        OK,
        MISSING
    }

    private final Path path;
    private final String name;
    private final String extension;
    private final long size;
    private final FileTime lastModified;
    private final String sourceDir;
    private boolean checked;
    private Status status;

    public FileItem(Path path,
                    String name,
                    String extension,
                    long size,
                    FileTime lastModified,
                    String sourceDir,
                    boolean checked,
                    Status status) {
        this.path = path;
        this.name = name;
        this.extension = extension;
        this.size = size;
        this.lastModified = lastModified;
        this.sourceDir = sourceDir;
        this.checked = checked;
        this.status = status;
    }

    public Path getPath() {
        return path;
    }

    public String getName() {
        return name;
    }

    public String getExtension() {
        return extension;
    }

    public FileType getFileType() {
        return FileType.fromExtension(extension);
    }

    public long getSize() {
        return size;
    }

    public FileTime getLastModified() {
        return lastModified;
    }

    public String getSourceDir() {
        return sourceDir;
    }

    public boolean isChecked() {
        return checked;
    }

    public void setChecked(boolean checked) {
        this.checked = checked;
    }

    public Status getStatus() {
        return status;
    }

    public void setStatus(Status status) {
        this.status = status;
    }

    public boolean isMissing() {
        return status == Status.MISSING;
    }
}
