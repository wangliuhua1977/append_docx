package app.docmerge;

import java.nio.file.Path;
import java.nio.file.attribute.FileTime;

public class FileItem {
    private final Path path;
    private final String name;
    private final String extension;
    private final long size;
    private final FileTime lastModified;

    public FileItem(Path path, String name, String extension, long size, FileTime lastModified) {
        this.path = path;
        this.name = name;
        this.extension = extension;
        this.size = size;
        this.lastModified = lastModified;
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

    public long getSize() {
        return size;
    }

    public FileTime getLastModified() {
        return lastModified;
    }
}
