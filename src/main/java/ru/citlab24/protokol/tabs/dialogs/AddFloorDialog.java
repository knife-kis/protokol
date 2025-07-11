package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Floor;

import javax.swing.*;
import java.awt.*;

public class AddFloorDialog extends JDialog {
    private boolean confirmed = false;
    private JTextField floorNumberField;
    private JComboBox<Floor.FloorType> floorTypeCombo;

    public AddFloorDialog(JFrame parent) {
        super(parent, "Добавить этаж", true);
        setSize(350, 200);
        setLocationRelativeTo(parent);
        initUI();
    }

    private void initUI() {
        JPanel mainPanel = new JPanel(new GridLayout(3, 2, 10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        // Поля ввода
        mainPanel.add(new JLabel("Название этажа:"));
        floorNumberField = new JTextField(); // Текстовое поле вместо спиннера
        mainPanel.add(floorNumberField);

        mainPanel.add(new JLabel("Тип этажа:"));
        floorTypeCombo = new JComboBox<>(Floor.FloorType.values());
        mainPanel.add(floorTypeCombo);

        // Кнопки
        JButton okButton = new JButton("Добавить");
        okButton.addActionListener(e -> {
            if (floorNumberField.getText().trim().isEmpty()) {
                JOptionPane.showMessageDialog(this, "Введите название этажа!", "Ошибка", JOptionPane.WARNING_MESSAGE);
                return;
            }
            confirmed = true;
            dispose();
        });

        JButton cancelButton = new JButton("Отмена");
        cancelButton.addActionListener(e -> dispose());

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(okButton);
        buttonPanel.add(cancelButton);

        add(mainPanel, BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);
    }

    public boolean showDialog() {
        setVisible(true);
        return confirmed;
    }

    public String getFloorNumber() { // Возвращаем строку вместо числа
        return floorNumberField.getText().trim();
    }

    public Floor.FloorType getFloorType() {
        return (Floor.FloorType) floorTypeCombo.getSelectedItem();
    }

    public void setFloorNumber(String number) {
        floorNumberField.setText(number);
    }

    public void setFloorType(String type) {
        floorTypeCombo.setSelectedItem(type);
    }
}