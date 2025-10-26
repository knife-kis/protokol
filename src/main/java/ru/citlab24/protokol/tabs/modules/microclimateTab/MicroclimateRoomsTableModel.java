package ru.citlab24.protokol.tabs.modules.microclimateTab;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.*;

final class MicroclimateRoomsTableModel extends AbstractTableModel {
    private final String[] NAMES = {"Измерения", "Комната", "Наружные стены"};
    private final Class<?>[] TYPES = {Boolean.class, String.class, Integer.class};

    private final Map<Integer, Boolean> selectionMap; // по id комнаты — ТОЛЬКО для микроклимата
    private final List<Room> rooms = new ArrayList<>();

    MicroclimateRoomsTableModel(Map<Integer, Boolean> selectionMap) {
        this.selectionMap = selectionMap;
    }

    void setRooms(List<Room> list) { rooms.clear(); rooms.addAll(list); fireTableDataChanged(); }
    void clear() { rooms.clear(); fireTableDataChanged(); }

    @Override public int getRowCount() { return rooms.size(); }
    @Override public int getColumnCount() { return NAMES.length; }
    @Override public String getColumnName(int c) { return NAMES[c]; }
    @Override public Class<?> getColumnClass(int c) { return TYPES[c]; }
    @Override public boolean isCellEditable(int row, int col) { return col != 1; }

    @Override public Object getValueAt(int row, int col) {
        Room r = rooms.get(row);
        return switch (col) {
            case 0 -> selectionMap.getOrDefault(r.getId(), r.isMicroclimateSelected());
            case 1 -> r.getName();
            case 2 -> r.getExternalWallsCount();
            default -> null;
        };
    }

    @Override public void setValueAt(Object aValue, int row, int col) {
        Room r = rooms.get(row);
        if (col == 0) {
            boolean v = (aValue instanceof Boolean) && (Boolean) aValue;
            selectionMap.put(r.getId(), v);
            r.setMicroclimateSelected(v);     // <<< ВАЖНО: НЕ setSelected()
        } else if (col == 2) {
            Integer val = (aValue instanceof Integer) ? (Integer) aValue : null;
            r.setExternalWallsCount(val);
        }
        fireTableRowsUpdated(row, row);
    }
}
