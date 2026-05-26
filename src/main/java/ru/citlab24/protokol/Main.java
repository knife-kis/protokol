package ru.citlab24.protokol;

import javax.swing.*;
import java.util.Locale;

public class Main {
    public static void main(String[] args) {
        Locale.setDefault(new Locale("ru", "RU"));
        SwingUtilities.invokeLater(() -> {
            SwingUtilities.invokeLater(() -> {
                AppTheme.install();
                javafx.application.Platform.setImplicitExit(false);

                MainFrame frame = new MainFrame();
                frame.setLocationByPlatform(true);
                frame.setVisible(true);
            });
        });
    }
}
