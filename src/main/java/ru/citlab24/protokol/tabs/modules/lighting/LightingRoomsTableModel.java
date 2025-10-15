package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.*;
import java.util.function.Consumer;

/** Таблица комнат: [чекбокс «Измерения» | «Комната»] */
final class LightingRoomsTableModel extends AbstractTableModel {

    private final String[] COLUMN_NAMES = {"Измерения", "Комната"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class};

    private final Map<Integer, Boolean> globalSelectionMap; // по id комнаты
    private final List<Room> rooms = new ArrayList<>();
    private final Consumer<Integer> onUserTouched; // уведомить вкладку, что юзер кликнул

    LightingRoomsTableModel(Map<Integer, Boolean> globalSelectionMap,
                            Consumer<Integer> onUserTouched) {
        this.globalSelectionMap = Objects.requireNonNull(globalSelectionMap);
        this.onUserTouched = Objects.requireNonNull(onUserTouched);
    }

    void setRooms(List<Room> newRooms) {
        rooms.clear();
        if (newRooms != null) rooms.addAll(newRooms);
        fireTableDataChanged();
    }

    Room getRoomAt(int row) {
        return (row >= 0 && row < rooms.size()) ? rooms.get(row) : null;
    }

    @Override public int getRowCount() { return rooms.size(); }
    @Override public int getColumnCount() { return COLUMN_NAMES.length; }
    @Override public String getColumnName(int column) { return COLUMN_NAMES[column]; }
    @Override public Class<?> getColumnClass(int columnIndex) { return COLUMN_TYPES[columnIndex]; }
    @Override public boolean isCellEditable(int rowIndex, int columnIndex) { return columnIndex == 0; }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Room r = rooms.get(rowIndex);
        if (columnIndex == 0) {
            Boolean v = globalSelectionMap.get(r.getId());
            return (v != null) ? v : Boolean.FALSE;
        }
        return r.getName();
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex != 0) return;
        Room r = rooms.get(rowIndex);
        boolean val = (aValue instanceof Boolean) && (Boolean) aValue;
        globalSelectionMap.put(r.getId(), val);
        onUserTouched.accept(r.getId());
        fireTableRowsUpdated(rowIndex, rowIndex);
    }
}
