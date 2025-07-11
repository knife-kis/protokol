package ru.citlab24.protokol.tabs.microclimateTab;

import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;

import javax.swing.*;
import java.awt.*;

public class MicroclimateTab extends JPanel {

    public MicroclimateTab() {
        setLayout(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        JPanel content = new JPanel(new GridLayout(0, 2, 10, 10));
        content.add(new JLabel("Температура (°C):"));
        content.add(new JTextField());
        content.add(new JLabel("Влажность (%):"));
        content.add(new JTextField());
        content.add(new JLabel("Давление (hPa):"));
        content.add(new JTextField());

        add(content, BorderLayout.NORTH);
        add(new JButton("Рассчитать параметры"), BorderLayout.SOUTH);
        JButton calculateButton = new JButton("Анализ микроклимата");
        calculateButton.setIcon(FontIcon.of(FontAwesomeSolid.CHART_LINE, 16, Color.WHITE));
        calculateButton.setBackground(new Color(103, 58, 183)); // Фиолетовый
        calculateButton.setForeground(Color.WHITE);
        calculateButton.setFocusPainted(false);
        calculateButton.setFont(new Font("Segoe UI", Font.BOLD, 14));

// Добавляем эффект при наведении
        calculateButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                calculateButton.setBackground(new Color(81, 45, 168));
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                calculateButton.setBackground(new Color(103, 58, 183));
            }
        });
    }
}