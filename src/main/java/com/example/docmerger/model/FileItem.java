package com.example.docmerger.model;

import java.nio.file.Path;
import java.nio.file.attribute.FileTime;
import java.util.Objects;

public final class FileItem {
    private final Path path;
    private final String fileName;
    private final long size;
    private final FileTime lastModified;
    private final String stableId;

    public FileItem(Path path, String fileName, long size, FileTime lastModified, String stableId) {
        this.path = Objects.requireNonNull(path, "path");
        this.fileName = Objects.requireNonNull(fileName, "fileName");
        this.size = size;
        this.lastModified = Objects.requireNonNull(lastModified, "lastModified");
        this.stableId = Objects.requireNonNull(stableId, "stableId");
    }

    public Path getPath() {
        return path;
    }

    public String getFileName() {
        return fileName;
    }

    public long getSize() {
        return size;
    }

    public FileTime getLastModified() {
        return lastModified;
    }

    public String getStableId() {
        return stableId;
    }
}
