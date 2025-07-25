package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class RadiationRoomsTableModel extends AbstractTableModel {
    private final String[] COLUMN_NAMES = {"", "Комната"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class};
    private final Map<Integer, Boolean> globalSelectionMap;
    private final Map<Integer, Boolean> selectionMap;
    private final List<Room> rooms = new ArrayList<>();
    private final Map<Room, Integer> roomToIdMap = new HashMap<>();

    public RadiationRoomsTableModel(Map<Integer, Boolean> globalSelectionMap) {
        this.globalSelectionMap = globalSelectionMap;
        this.selectionMap = globalSelectionMap;
    }

    // Добавлены недостающие методы
    public void addRoom(Room room, int roomId) {
        rooms.add(room);
        roomToIdMap.put(room, roomId);
        fireTableRowsInserted(rooms.size()-1, rooms.size()-1);
    }

    public void clear() {
        int size = rooms.size();
        rooms.clear();
        roomToIdMap.clear();
        if (size > 0) {
            fireTableRowsDeleted(0, size - 1);
        }
    }

    public Room getRoomAt(int rowIndex) {
        return rooms.get(rowIndex);
    }

    @Override
    public int getRowCount() {
        return rooms.size();
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
        return COLUMN_TYPES[columnIndex];
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Room room = rooms.get(rowIndex);
        int roomId = roomToIdMap.get(room);

        if (columnIndex == 0) {
            return selectionMap.get(roomId);
        } else {
            return room.getName(); // Или другое свойство комнаты
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex == 0) {
            Room room = rooms.get(rowIndex);
            int roomId = roomToIdMap.get(room);
            selectionMap.put(roomId, (Boolean) aValue);
            fireTableCellUpdated(rowIndex, columnIndex);
        }
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return columnIndex == 0;
    }
}