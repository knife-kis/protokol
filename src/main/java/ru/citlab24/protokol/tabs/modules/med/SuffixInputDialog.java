package ru.citlab24.protokol.tabs.modules.med;

import javax.swing.*;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;

class SuffixInputDialog extends JDialog {
    private boolean confirmed = false;
    private final List<JTextField> suffixFields = new ArrayList<>();

    public SuffixInputDialog(Frame owner, String roomName, int suffixCount) {
        super(owner, "Ввод суффиксов для комнаты: " + roomName, true);
        setLayout(new GridLayout(suffixCount + 1, 1)); // +1 для кнопок
        setSize(400, 50 + suffixCount * 40);

        // Создаем поля для ввода суффиксов
        for (int i = 0; i < suffixCount; i++) {
            JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT));
            panel.add(new JLabel("Суффикс " + (i + 1) + ": "));
            JTextField textField = new JTextField(10);
            suffixFields.add(textField);
            panel.add(textField);
            add(panel);
        }

        JPanel buttonPanel = new JPanel();
        JButton btnConfirm = new JButton("Подтвердить");
        JButton btnCancel = new JButton("Отмена");

        btnConfirm.addActionListener(e -> {
            confirmed = true;
            dispose();
        });
        btnCancel.addActionListener(e -> dispose());

        buttonPanel.add(btnConfirm);
        buttonPanel.add(btnCancel);
        add(buttonPanel);
    }

    public boolean isConfirmed() {
        return confirmed;
    }

    public List<String> getSuffixes() {
        List<String> suffixes = new ArrayList<>();
        for (JTextField field : suffixFields) {
            suffixes.add(field.getText().trim());
        }
        return suffixes;
    }
}