package com.example.docmerger.service;

import com.example.docmerger.model.FileItem;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public final class FileScanner {
    public List<FileItem> scan(Path directory) throws IOException {
        List<FileItem> items = new ArrayList<>();
        try (var stream = Files.list(directory)) {
            stream.filter(Files::isRegularFile)
                    .filter(path -> isDocFile(path.getFileName().toString()))
                    .forEach(path -> {
                        try {
                            BasicFileAttributes attrs = Files.readAttributes(path, BasicFileAttributes.class);
                            String stableId = path.toAbsolutePath().normalize().toString();
                            items.add(new FileItem(
                                    path,
                                    path.getFileName().toString(),
                                    attrs.size(),
                                    attrs.lastModifiedTime(),
                                    stableId
                            ));
                        } catch (IOException ignored) {
                            // Skip unreadable files silently.
                        }
                    });
        }
        return items;
    }

    private boolean isDocFile(String fileName) {
        String lower = fileName.toLowerCase(Locale.ROOT);
        return lower.endsWith(".doc") || lower.endsWith(".docx");
    }
}
