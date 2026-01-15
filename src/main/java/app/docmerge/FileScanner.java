package app.docmerge;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class FileScanner {
    private static final Pattern LEADING_NUMBER = Pattern.compile("^(\\d+)");

    public List<FileItem> scan(Path directory) throws IOException {
        List<FileItem> items = new ArrayList<>();
        if (directory == null || !Files.isDirectory(directory)) {
            return items;
        }
        try (var stream = Files.list(directory)) {
            stream.filter(path -> Files.isRegularFile(path))
                    .filter(path -> isDocFile(path.getFileName().toString()))
                    .forEach(path -> {
                        try {
                            BasicFileAttributes attrs = Files.readAttributes(path, BasicFileAttributes.class);
                            String fileName = path.getFileName().toString();
                            String extension = extensionOf(fileName);
                            items.add(new FileItem(path, fileName, extension, attrs.size(), attrs.lastModifiedTime()));
                        } catch (IOException ignored) {
                            // 忽略无法读取的文件
                        }
                    });
        }
        items.sort(defaultComparator());
        return items;
    }

    public Comparator<FileItem> defaultComparator() {
        return Comparator.comparingInt((FileItem item) -> 0)
                .thenComparing((FileItem a, FileItem b) -> {
                    String nameA = a.getName();
                    String nameB = b.getName();
                    Integer numA = leadingNumber(nameA);
                    Integer numB = leadingNumber(nameB);
                    if (numA != null && numB != null && !numA.equals(numB)) {
                        return Integer.compare(numA, numB);
                    }
                    if (numA != null && numB == null) {
                        return -1;
                    }
                    if (numA == null && numB != null) {
                        return 1;
                    }
                    return nameA.toLowerCase(Locale.ROOT).compareTo(nameB.toLowerCase(Locale.ROOT));
                });
    }

    private boolean isDocFile(String name) {
        String lower = name.toLowerCase(Locale.ROOT);
        return lower.endsWith(".docx") || lower.endsWith(".doc");
    }

    private String extensionOf(String name) {
        int idx = name.lastIndexOf('.');
        return idx >= 0 ? name.substring(idx + 1) : "";
    }

    private Integer leadingNumber(String name) {
        Matcher matcher = LEADING_NUMBER.matcher(name);
        if (matcher.find()) {
            try {
                return Integer.parseInt(matcher.group(1));
            } catch (NumberFormatException ignored) {
                return null;
            }
        }
        return null;
    }
}
