package ru.citlab24.protokol;

import com.formdev.flatlaf.FlatLightLaf;
import javax.swing.*;
public class Main {
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            // Устанавливаем современную тему
            FlatLightLaf.setup();

            // Настройка глобальных параметров
            UIManager.put("Button.arc", 8);
            UIManager.put("Component.arc", 8);
            UIManager.put("TextComponent.arc", 5);
            UIManager.put("ProgressBar.arc", 8);
            UIManager.put("ScrollBar.width", 14);

            MainFrame frame = new MainFrame();
            frame.setVisible(true);
        });
    }
}