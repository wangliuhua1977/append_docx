package com.example.docmerger;

import com.example.docmerger.ui.DocMergerFrame;

import javax.swing.SwingUtilities;

public final class Main {
    private Main() {
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            DocMergerFrame frame = new DocMergerFrame();
            frame.setVisible(true);
        });
    }
}
