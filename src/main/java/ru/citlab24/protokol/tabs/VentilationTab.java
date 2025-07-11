package ru.citlab24.protokol.tabs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public class VentilationTab extends JPanel {
    private final Building building;
    private final VentilationTableModel tableModel = new VentilationTableModel();
    private final JTable ventilationTable = new JTable(tableModel);

    private static final List<String> TARGET_ROOMS = List.of(
            "кухня", "кухня-ниша", "санузел", "ванная", "совмещенный"
    );

    private static final List<String> TARGET_FLOORS = List.of(
            "жилой", "смешанный"
    );

    public VentilationTab(Building building) {
        this.building = building;
        initUI();
        loadVentilationData();
    }

    private void initUI() {
        setLayout(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        // Настройка таблицы
        ventilationTable.setRowHeight(30);
        ventilationTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        ventilationTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));

        // Редакторы ячеек
        ventilationTable.getColumnModel().getColumn(3).setCellEditor(
                new SpinnerEditor(1, 1, 10));
        ventilationTable.getColumnModel().getColumn(4).setCellEditor(
                new SpinnerEditor(0.008, 0.001, 0.1, 0.001));

        add(new JScrollPane(ventilationTable), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    private JPanel createButtonPanel() {
        JButton saveBtn = new JButton("Сохранить расчет");
        saveBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        saveBtn.addActionListener(e -> saveCalculations());

        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        panel.add(saveBtn);
        return panel;
    }

    public void refreshData() {
        loadVentilationData();
    }

    private void loadVentilationData() {
        tableModel.clearData();

        for (Floor floor : building.getFloors()) {
            String floorType = floor.getType().name().toLowerCase(Locale.ROOT);

            if (!containsAny(floorType, TARGET_FLOORS)) continue;

            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    String roomName = room.getName().toLowerCase(Locale.ROOT);

                    if (!containsAny(roomName, TARGET_ROOMS)) continue;

                    tableModel.addRecord(new VentilationRecord(
                            floor.getNumber(),
                            space.getIdentifier(),
                            room.getName(),
                            room.getVentilationChannels(),
                            room.getVentilationSectionArea(),
                            room.getVolume(),
                            room
                    ));
                }
            }
        }
        tableModel.fireTableDataChanged();
    }

    private boolean containsAny(String source, List<String> targets) {
        return targets.stream().anyMatch(source::contains);
    }

    private void saveCalculations() {
        // Сохранение данных обратно в модель
        for (VentilationRecord record : tableModel.getRecords()) {
            record.roomRef().setVentilationChannels(record.channels());
            record.roomRef().setVentilationSectionArea(record.sectionArea());
        }
        JOptionPane.showMessageDialog(this, "Расчеты сохранены успешно!", "Сохранение", JOptionPane.INFORMATION_MESSAGE);
    }

    // Модель таблицы
    private static class VentilationTableModel extends AbstractTableModel {
        private final String[] COLUMN_NAMES = {
                "Этаж", "Помещение", "Комната",
                "Кол-во каналов", "Сечение (кв.м)", "Объем (куб.м)"
        };

        private final List<VentilationRecord> records = new ArrayList<>();

        public void clearData() {
            records.clear();
        }

        public void addRecord(VentilationRecord record) {
            records.add(record);
        }

        public List<VentilationRecord> getRecords() {
            return records;
        }

        @Override
        public int getRowCount() {
            return records.size();
        }

        @Override
        public int getColumnCount() {
            return COLUMN_NAMES.length;
        }

        @Override
        public String getColumnName(int column) {
            return COLUMN_NAMES[column];
        }

        @Override
        public Class<?> getColumnClass(int columnIndex) {
            return switch (columnIndex) {
                case 0, 1, 2 -> String.class;
                case 3 -> Integer.class;
                case 4, 5 -> Double.class;
                default -> Object.class;
            };
        }

        @Override
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            return columnIndex == 3 || columnIndex == 4;
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            VentilationRecord record = records.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> record.floor();
                case 1 -> record.space();
                case 2 -> record.room();
                case 3 -> record.channels();
                case 4 -> record.sectionArea();
                case 5 -> record.volume();
                default -> null;
            };
        }

        @Override
        public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
            VentilationRecord record = records.get(rowIndex);
            records.set(rowIndex, switch (columnIndex) {
                case 3 -> record.withChannels((Integer) aValue);
                case 4 -> record.withSectionArea((Double) aValue);
                default -> record;
            });
            fireTableCellUpdated(rowIndex, columnIndex);
        }
    }

    // Record с ссылкой на модель
    private record VentilationRecord(
            String floor,
            String space,
            String room,
            int channels,
            double sectionArea,
            double volume,
            Room roomRef
    ) {
        public VentilationRecord withChannels(int newChannels) {
            return new VentilationRecord(floor, space, room, newChannels, sectionArea, volume, roomRef);
        }

        public VentilationRecord withSectionArea(double newArea) {
            return new VentilationRecord(floor, space, room, channels, newArea, volume, roomRef);
        }
    }
}