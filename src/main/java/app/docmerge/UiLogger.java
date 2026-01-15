package app.docmerge;

import javax.swing.JTextArea;
import javax.swing.SwingUtilities;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class UiLogger {
    private static final DateTimeFormatter FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");
    private final JTextArea textArea;

    public UiLogger(JTextArea textArea) {
        this.textArea = textArea;
    }

    public void info(String message) {
        append("[信息] " + message);
    }

    public void warn(String message) {
        append("[警告] " + message);
    }

    public void error(String message, Throwable throwable) {
        StringWriter writer = new StringWriter();
        if (throwable != null) {
            throwable.printStackTrace(new PrintWriter(writer));
        }
        append("[错误] " + message + (throwable == null ? "" : "\n" + writer));
    }

    public void clear() {
        SwingUtilities.invokeLater(() -> textArea.setText(""));
    }

    private void append(String message) {
        String time = LocalDateTime.now().format(FORMATTER);
        SwingUtilities.invokeLater(() -> {
            textArea.append("[" + time + "] " + message + "\n");
            textArea.setCaretPosition(textArea.getDocument().getLength());
        });
    }
}
