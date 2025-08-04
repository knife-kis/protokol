package ru.citlab24.protokol.tabs.modules.med;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ArrayList;
import java.util.List; // Исправленный импорт

class SplitRoomDialog extends JDialog {
    private boolean confirmed = false;
    private final DefaultTableModel tableModel;
    private final JTable table;
    private final JButton btnConfirm;
    private final JButton btnCancel;
    private final JButton btnAddRow;

    public SplitRoomDialog(Frame owner, String roomName) {
        super(owner, "Разделение комнаты: " + roomName, true);
        setSize(400, 300);
        setLayout(new BorderLayout());

        // Таблица для ввода суффиксов
        tableModel = new DefaultTableModel(new String[]{"Суффикс"}, 0);
        table = new JTable(tableModel);
        add(new JScrollPane(table), BorderLayout.CENTER);

        // Панель кнопок
        JPanel buttonPanel = new JPanel();
        btnAddRow = new JButton("Добавить строку");
        btnConfirm = new JButton("Подтвердить");
        btnCancel = new JButton("Отмена");

        btnAddRow.addActionListener(e -> tableModel.addRow(new Object[]{""}));
        btnConfirm.addActionListener(e -> {
            confirmed = true;
            dispose();
        });
        btnCancel.addActionListener(e -> dispose());

        buttonPanel.add(btnAddRow);
        buttonPanel.add(btnConfirm);
        buttonPanel.add(btnCancel);
        add(buttonPanel, BorderLayout.SOUTH);
    }

    public boolean isConfirmed() {
        return confirmed;
    }

    public List<String> getSuffixes() {
        List<String> suffixes = new ArrayList<>();
        for (int i = 0; i < tableModel.getRowCount(); i++) {
            String suffix = (String) tableModel.getValueAt(i, 0);
            if (suffix != null && !suffix.trim().isEmpty()) {
                suffixes.add(suffix.trim());
            }
        }
        return suffixes;
    }
}
