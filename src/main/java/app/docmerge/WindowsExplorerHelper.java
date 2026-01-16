package app.docmerge;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

public final class WindowsExplorerHelper {
    private static final int MAX_RETRY = 5;
    private static final long RETRY_INTERVAL_MS = 40L;

    private WindowsExplorerHelper() {
    }

    public static void openAndSelect(Path file, UiLogger logger) {
        if (file == null) {
            return;
        }
        if (!isWindows()) {
            return;
        }
        Path target = file.toAbsolutePath();
        if (!waitForFile(target, logger)) {
            return;
        }
        try {
            String argument = "/select," + target;
            new ProcessBuilder("explorer.exe", argument).start();
        } catch (IOException e) {
            logger.warn("打开资源管理器失败：" + target + "，原因：" + e.getMessage());
        }
    }

    private static boolean waitForFile(Path file, UiLogger logger) {
        for (int i = 0; i < MAX_RETRY; i++) {
            if (Files.exists(file)) {
                return true;
            }
            try {
                Thread.sleep(RETRY_INTERVAL_MS);
            } catch (InterruptedException ignored) {
                Thread.currentThread().interrupt();
                break;
            }
        }
        logger.warn("输出文件不存在，无法打开资源管理器：" + file);
        return false;
    }

    private static boolean isWindows() {
        String osName = System.getProperty("os.name");
        return osName != null && osName.toLowerCase().contains("windows");
    }
}
