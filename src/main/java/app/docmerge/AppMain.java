package app.docmerge;

import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.SwingUtilities;

public class AppMain {
    public static void main(String[] args) {
        FlatLightLaf.setup();
        SwingUtilities.invokeLater(() -> {
            MainFrame frame = new MainFrame();
            frame.setVisible(true);
        });
    }
}
