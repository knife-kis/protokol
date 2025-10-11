package ru.citlab24.protokol;

import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
import java.util.Locale;

public class Main {
    public static void main(String[] args) {
        Locale.setDefault(new Locale("ru", "RU"));
        SwingUtilities.invokeLater(() -> {
            SwingUtilities.invokeLater(() -> {
                // Гладкие шрифты + современные оконные декорации
                System.setProperty("awt.useSystemAAFontSettings", "on");
                System.setProperty("swing.aatext", "true");
                System.setProperty("flatlaf.useWindowDecorations", "true");
                com.formdev.flatlaf.FlatLaf.setUseNativeWindowDecorations(true);

                // Светлая по умолчанию (можно FlatDarkLaf / FlatIntelliJLaf / FlatDarculaLaf)
                UIManager.put("Component.arc", 12);
                UIManager.put("Button.arc", 12);
                UIManager.put("TextComponent.arc", 12);
                UIManager.put("ScrollBar.width", 14);
                UIManager.put("ScrollBar.thumbArc", 999); // «пилюля»
                UIManager.put("Component.focusWidth", 1);

                try {
                    UIManager.setLookAndFeel(new com.formdev.flatlaf.FlatIntelliJLaf());
                    com.formdev.flatlaf.FlatLaf.updateUI();
                } catch (Exception ignore) {}

                MainFrame frame = new MainFrame();
                frame.setLocationByPlatform(true);
                frame.setVisible(true);
            });
        });
    }
}