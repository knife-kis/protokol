package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Floor;

import javax.swing.*;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;

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

// Enter в поле ввода = нажать «Добавить»
        floorNumberField.addActionListener(e -> okButton.doClick());

        JButton cancelButton = new JButton("Отмена");
        cancelButton.addActionListener(e -> dispose());

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(okButton);
        buttonPanel.add(cancelButton);

        add(mainPanel, BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);

// Синяя «дефолтная» кнопка + реакция на Enter по всему диалогу
        getRootPane().setDefaultButton(okButton);

// Фокус сразу в поле (как в диалоге комнаты)
        addWindowListener(new WindowAdapter() {
            @Override public void windowOpened(WindowEvent e) {
                floorNumberField.requestFocusInWindow();
                floorNumberField.selectAll();
            }
        });
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
    public void setFloorType(Floor.FloorType type) {
        floorTypeCombo.setSelectedItem(type);
    }

    public void setFloorNumber(String number) {
        floorNumberField.setText(number);
    }

}