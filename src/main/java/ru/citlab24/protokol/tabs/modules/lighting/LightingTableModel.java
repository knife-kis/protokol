package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/** Таблица комнат как в радиации: [Измерения],[Комната] */
class LightingRoomsTableModel extends AbstractTableModel {
    private final String[] COLUMN_NAMES = {"Измерения", "Комната"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class};

    private final Map<Integer, Boolean> globalSelectionMap;
    private final List<Room> rooms = new ArrayList<>();

    LightingRoomsTableModel(Map<Integer, Boolean> globalSelectionMap) {
        this.globalSelectionMap = globalSelectionMap;
    }

    public void setRooms(List<Room> list) {
        rooms.clear();
        if (list != null) rooms.addAll(list);
        fireTableDataChanged();
    }

    public Room getRoomAt(int row) {
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
        return switch (columnIndex) {
            case 0 -> Boolean.TRUE.equals(globalSelectionMap.getOrDefault(r.getId(), Boolean.FALSE));
            case 1 -> r.getName();
            default -> null;
        };
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex != 0) return;
        Room r = rooms.get(rowIndex);
        boolean val = (aValue instanceof Boolean) && (Boolean) aValue;
        globalSelectionMap.put(r.getId(), val);
        fireTableRowsUpdated(rowIndex, rowIndex);
    }
}
