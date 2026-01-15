package app.docmerge;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

public class LibreOfficeConverter {
    public Optional<Path> findSofficeExecutable() {
        String pathEnv = System.getenv("PATH");
        if (pathEnv != null) {
            for (String part : pathEnv.split(System.getProperty("path.separator"))) {
                Path candidate = Paths.get(part, isWindows() ? "soffice.exe" : "soffice");
                if (Files.isExecutable(candidate)) {
                    return Optional.of(candidate);
                }
            }
        }
        if (isWindows()) {
            List<Path> candidates = new ArrayList<>();
            String programFiles = System.getenv("ProgramFiles");
            String programFilesX86 = System.getenv("ProgramFiles(x86)");
            if (programFiles != null) {
                candidates.add(Path.of(programFiles, "LibreOffice", "program", "soffice.exe"));
            }
            if (programFilesX86 != null) {
                candidates.add(Path.of(programFilesX86, "LibreOffice", "program", "soffice.exe"));
            }
            for (Path candidate : candidates) {
                if (Files.isExecutable(candidate)) {
                    return Optional.of(candidate);
                }
            }
        }
        return Optional.empty();
    }

    public Optional<Path> convertToDocx(Path inputDoc, Path tempDir, UiLogger logger) {
        Optional<Path> soffice = findSofficeExecutable();
        if (soffice.isEmpty()) {
            return Optional.empty();
        }
        try {
            Files.createDirectories(tempDir);
        } catch (IOException e) {
            logger.error("无法创建转换临时目录：" + tempDir, e);
            return Optional.empty();
        }

        List<String> command = new ArrayList<>();
        command.add(soffice.get().toString());
        command.add("--headless");
        command.add("--nologo");
        command.add("--convert-to");
        command.add("docx");
        command.add("--outdir");
        command.add(tempDir.toString());
        command.add(inputDoc.toString());

        ProcessBuilder builder = new ProcessBuilder(command);
        builder.redirectErrorStream(true);
        try {
            Process process = builder.start();
            int code = process.waitFor();
            if (code != 0) {
                logger.warn("LibreOffice 转换失败，退出码：" + code);
                return Optional.empty();
            }
            String baseName = inputDoc.getFileName().toString();
            int idx = baseName.lastIndexOf('.');
            String targetName = idx >= 0 ? baseName.substring(0, idx) + ".docx" : baseName + ".docx";
            Path output = tempDir.resolve(targetName);
            if (Files.exists(output)) {
                return Optional.of(output);
            }
            logger.warn("LibreOffice 未生成输出文件：" + output);
            return Optional.empty();
        } catch (IOException e) {
            logger.error("LibreOffice 转换异常", e);
            return Optional.empty();
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
            logger.error("LibreOffice 转换被中断", e);
            return Optional.empty();
        }
    }

    private boolean isWindows() {
        return System.getProperty("os.name").toLowerCase(Locale.ROOT).contains("win");
    }
}
