package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.*;

/** Таблица для вкладки «Освещение»: [чекбокс «Измерения» | «Комната»]. */
final class ArtificialLightingRoomsTableModel extends AbstractTableModel {

    private final String[] COLUMN_NAMES = {"Измерения", "Комната"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class};

    private final Map<Integer, Boolean> selectionMap; // roomId → selected
    private final List<Room> rooms = new ArrayList<>();

    ArtificialLightingRoomsTableModel(Map<Integer, Boolean> selectionMap) {
        this.selectionMap = selectionMap;
    }

    public void setRooms(List<Room> newRooms) {
        rooms.clear();
        if (newRooms != null) rooms.addAll(newRooms);
        fireTableDataChanged();
    }

    @Override public int getRowCount() { return rooms.size(); }
    @Override public int getColumnCount() { return COLUMN_NAMES.length; }
    @Override public String getColumnName(int column) { return COLUMN_NAMES[column]; }
    @Override public Class<?> getColumnClass(int columnIndex) { return COLUMN_TYPES[columnIndex]; }
    @Override public boolean isCellEditable(int rowIndex, int columnIndex) { return columnIndex == 0; }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Room r = rooms.get(rowIndex);
        return (columnIndex == 0)
                ? selectionMap.getOrDefault(r.getId(), Boolean.TRUE) // дефолт true для видимых комнат
                : r.getName();
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex != 0) return;
        Room r = rooms.get(rowIndex);
        boolean v = (aValue instanceof Boolean) ? (Boolean) aValue : false;
        selectionMap.put(r.getId(), v);
        fireTableCellUpdated(rowIndex, columnIndex);
    }
    public Room getRoomAt(int row) {
        return (row >= 0 && row < rooms.size()) ? rooms.get(row) : null;
    }

}
