package ru.citlab24.protokol.tabs;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;
import ru.citlab24.protokol.tabs.VentilationRecord;

public class VentilationTableModel extends AbstractTableModel {
    private final String[] COLUMN_NAMES = {
            "Этаж", "Помещение", "Комната",
            "Кол-во каналов", "Сечение (кв.м)", "Объем (куб.м)"
    };

    private final VentilationTab ventilationTab;
    private final List<VentilationRecord> records = new ArrayList<>();

    public VentilationTableModel(VentilationTab ventilationTab) {
        this.ventilationTab = ventilationTab;
    }

    public void clearData() {
        records.clear();
    }

    public void addRecord(VentilationRecord record) {
        records.add(record);
    }

    public List<VentilationRecord> getRecords() {
        return new ArrayList<>(records);
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
            case 4 -> Double.class;
            case 5 -> Object.class;
            default -> Object.class;
        };
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return columnIndex == 3 || columnIndex == 4 || columnIndex == 5;
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
        if (aValue == null) return;

        VentilationRecord record = records.get(rowIndex);
        switch (columnIndex) {
            case 3 -> handleChannelsUpdate(aValue, rowIndex, record);
            case 4 -> handleSectionUpdate(aValue, rowIndex, record);
            case 5 -> updateVolume(aValue, rowIndex, record);
        }
    }

    private void handleChannelsUpdate(Object aValue, int rowIndex, VentilationRecord record) {
        if (!(aValue instanceof Number numberValue)) return;

        int newValue = numberValue.intValue();
        int oldValue = record.channels();
        if (newValue == oldValue) return;

        String category = ventilationTab.getRoomCategory(record.room());
        if (category == null) {
            updateRecordAndRoom(record, rowIndex, newValue);
            return;
        }

        int option = showConfirmationDialog(
                "Вы хотите изменить количество каналов во всех комнатах типа '" + category + "'?"
        );

        switch (option) {
            case 0 -> updateAllRoomsOfType(category, newValue);
            case 1 -> updateRecordAndRoom(record, rowIndex, newValue);
            case 2 -> fireTableCellUpdated(rowIndex, 3);
        }
    }

    private void handleSectionUpdate(Object aValue, int rowIndex, VentilationRecord record) {
        if (!(aValue instanceof Number numberValue)) return;

        double newValue = numberValue.doubleValue();
        double oldValue = record.sectionArea();
        if (Math.abs(newValue - oldValue) < 0.000001) return;

        String category = ventilationTab.getRoomCategory(record.room());
        if (category == null) {
            updateRecordAndSection(record, rowIndex, newValue);
            return;
        }

        int option = showConfirmationDialog(
                "Вы хотите изменить сечение каналов во всех комнатах типа '" + category + "'?"
        );

        switch (option) {
            case 0 -> updateAllRoomsOfTypeForSection(category, newValue);
            case 1 -> updateRecordAndSection(record, rowIndex, newValue);
            case 2 -> fireTableCellUpdated(rowIndex, 4);
        }
    }

    private void updateVolume(Object aValue, int rowIndex, VentilationRecord record) {
        if (aValue instanceof Number numberValue) {
            double doubleValue = numberValue.doubleValue();
            records.set(rowIndex, record.withVolume(doubleValue));
            fireTableCellUpdated(rowIndex, 5);
        }
    }

    private int showConfirmationDialog(String message) {
        return JOptionPane.showOptionDialog(
                ventilationTab,
                message,
                "Подтверждение",
                JOptionPane.YES_NO_CANCEL_OPTION,
                JOptionPane.QUESTION_MESSAGE,
                null,
                new String[]{"Да", "Нет", "Отмена"},
                "Нет"
        );
    }

    private void updateAllRoomsOfTypeForSection(String category, double newValue) {
        ventilationTab.getBuilding().getFloors().forEach(floor ->
                floor.getSpaces().forEach(space ->
                        space.getRooms().stream()
                                .filter(room -> category.equals(ventilationTab.getRoomCategory(room.getName())))
                                .forEach(room -> room.setVentilationSectionArea(newValue))
                )
        );

        for (int i = 0; i < records.size(); i++) {
            VentilationRecord r = records.get(i);
            if (category.equals(ventilationTab.getRoomCategory(r.room()))) {
                records.set(i, r.withSectionArea(newValue));
                fireTableCellUpdated(i, 4);
            }
        }
    }

    private void updateRecordAndSection(VentilationRecord record, int rowIndex, double newValue) {
        records.set(rowIndex, record.withSectionArea(newValue));
        record.roomRef().setVentilationSectionArea(newValue);
        fireTableCellUpdated(rowIndex, 4);
    }

    private void updateAllRoomsOfType(String category, int newValue) {
        ventilationTab.getBuilding().getFloors().forEach(floor ->
                floor.getSpaces().forEach(space ->
                        space.getRooms().stream()
                                .filter(room -> category.equals(ventilationTab.getRoomCategory(room.getName())))
                                .forEach(room -> room.setVentilationChannels(newValue))
                )
        );

        for (int i = 0; i < records.size(); i++) {
            VentilationRecord r = records.get(i);
            if (category.equals(ventilationTab.getRoomCategory(r.room()))) {
                records.set(i, r.withChannels(newValue));
                fireTableCellUpdated(i, 3);
            }
        }
    }

    private void updateRecordAndRoom(VentilationRecord record, int rowIndex, int newValue) {
        records.set(rowIndex, record.withChannels(newValue));
        record.roomRef().setVentilationChannels(newValue);
        fireTableCellUpdated(rowIndex, 3);
    }
}