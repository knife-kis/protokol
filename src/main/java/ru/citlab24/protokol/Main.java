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
                    java.awt.Font baseFont = new java.awt.Font("Segoe UI", java.awt.Font.PLAIN, 15);
                    javax.swing.plaf.FontUIResource f = new javax.swing.plaf.FontUIResource(baseFont);
                    java.util.Enumeration<?> keys = javax.swing.UIManager.getDefaults().keys();
                    while (keys.hasMoreElements()) {
                        Object key = keys.nextElement();
                        Object val = javax.swing.UIManager.get(key);
                        if (val instanceof javax.swing.plaf.FontUIResource) {
                            javax.swing.UIManager.put(key, f);
                        }
                    }

// Дальше — ЛАФ
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