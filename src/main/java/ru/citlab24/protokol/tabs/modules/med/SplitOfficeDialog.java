package ru.citlab24.protokol.tabs.modules.med;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;

public class SplitOfficeDialog extends JDialog {
    private boolean confirmed = false;
    private final List<RoomData> roomDataList = new ArrayList<>();
    private final DefaultTableModel tableModel;

    public SplitOfficeDialog(Frame owner, String spaceName) {
        super(owner, "Разделить офис на точки", true);
        setLayout(new BorderLayout(10, 10));
        setSize(500, 400);
        setLocationRelativeTo(owner);

        // Панель ввода количества точек
        JPanel inputPanel = new JPanel();
        inputPanel.add(new JLabel("Количество точек:"));
        JSpinner countSpinner = new JSpinner(new SpinnerNumberModel(2, 1, 10, 1));
        inputPanel.add(countSpinner);

        JButton createButton = new JButton("Создать");
        inputPanel.add(createButton);
        add(inputPanel, BorderLayout.NORTH);

        // Таблица для ввода данных
        tableModel = new DefaultTableModel(new Object[]{"Идентификатор помещения", "Название комнаты"}, 0);
        JTable table = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(table);
        add(scrollPane, BorderLayout.CENTER);

        // Кнопки подтверждения
        JPanel buttonPanel = new JPanel();
        JButton okButton = new JButton("OK");
        JButton cancelButton = new JButton("Отмена");
        buttonPanel.add(okButton);
        buttonPanel.add(cancelButton);
        add(buttonPanel, BorderLayout.SOUTH);

        // Генерация полей при нажатии "Создать"
        createButton.addActionListener(e -> {
            int count = (Integer) countSpinner.getValue();
            tableModel.setRowCount(0);
            for (int i = 1; i <= count; i++) {
                tableModel.addRow(new Object[]{
                        spaceName + " в осях ?-?",
                        "Комната " + i
                });
            }
        });

        // Обработка OK
        okButton.addActionListener(e -> {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                String spaceId = (String) tableModel.getValueAt(i, 0);
                String roomName = (String) tableModel.getValueAt(i, 1);

                if (spaceId.isBlank() || roomName.isBlank()) {
                    JOptionPane.showMessageDialog(this, "Заполните все поля", "Ошибка", JOptionPane.ERROR_MESSAGE);
                    return;
                }
                roomDataList.add(new RoomData(spaceId, roomName));
            }
            confirmed = true;
            dispose();
        });

        cancelButton.addActionListener(e -> dispose());
    }

    public boolean isConfirmed() {
        return confirmed;
    }

    public List<RoomData> getRoomDataList() {
        return roomDataList;
    }

    public static class RoomData {
        private final String spaceIdentifier;
        private final String roomName;

        public RoomData(String spaceIdentifier, String roomName) {
            this.spaceIdentifier = spaceIdentifier;
            this.roomName = roomName;
        }

        // Геттеры
        public String getSpaceIdentifier() { return spaceIdentifier; }
        public String getRoomName() { return roomName; }
    }
}