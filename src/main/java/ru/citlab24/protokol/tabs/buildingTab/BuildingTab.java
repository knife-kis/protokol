package ru.citlab24.protokol.tabs.buildingTab;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;

public class BuildingTab extends JPanel {

    public BuildingTab() {
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));
        initComponents();
    }

    private void initComponents() {
        // Панель ввода с современным заголовком
        JPanel inputPanel = new JPanel(new GridLayout(0, 2, 15, 10));
        inputPanel.setBorder(BorderFactory.createTitledBorder(
                "Основные параметры здания"));

        // Поля с подсказками
        addField(inputPanel, "Название здания:", "Офисный комплекс 'Северный'");
        addField(inputPanel, "Площадь (м²):", "1500");
        addField(inputPanel, "Высота (м):", "24.5");
        addField(inputPanel, "Тип конструкции:", "Монолитный каркас");
        addField(inputPanel, "Год постройки:", "2020");
        addField(inputPanel, "Количество этажей:", "8");

        // Расширенные параметры
        JPanel advancedPanel = new JPanel(new BorderLayout());
        advancedPanel.setBorder(BorderFactory.createTitledBorder("Дополнительные характеристики"));

        JTextArea notesArea = new JTextArea(5, 20);
        notesArea.setLineWrap(true);
        notesArea.setWrapStyleWord(true);

        JButton calcButton = new JButton("Рассчитать показатели");
        calcButton.setIcon(FontIcon.of(FontAwesomeSolid.CALCULATOR, 16, Color.WHITE));
        calcButton.setBackground(new Color(46, 125, 50));
        calcButton.setForeground(Color.WHITE);
        calcButton.setFocusPainted(false);
        calcButton.addActionListener(this::calculate);

        advancedPanel.add(new JScrollPane(notesArea), BorderLayout.CENTER);
        advancedPanel.add(calcButton, BorderLayout.SOUTH);

        // Компоновка
        add(inputPanel, BorderLayout.NORTH);
        add(advancedPanel, BorderLayout.CENTER);
    }

    private void addField(JPanel panel, String label, String placeholder) {
        JLabel lbl = new JLabel(label);
        lbl.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        panel.add(lbl);

        JTextField field = new JTextField();
        field.putClientProperty("JTextField.placeholderText", placeholder);
        panel.add(field);
    }

    private void calculate(ActionEvent e) {
        JOptionPane.showMessageDialog(this,
                "Расчет выполнен успешно!",
                "Результаты",
                JOptionPane.INFORMATION_MESSAGE);
    }
}