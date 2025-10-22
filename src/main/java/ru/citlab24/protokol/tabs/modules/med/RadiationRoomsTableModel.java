package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

class RadiationRoomsTableModel extends AbstractTableModel {
    private final String[] COLUMN_NAMES = {"Измерения", "Комната"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class};
    private final Map<Integer, Boolean> globalSelectionMap;
    private final RadiationTab radiationTab;
    private final List<Room> rooms = new ArrayList<>();

    public RadiationRoomsTableModel(Map<Integer, Boolean> globalSelectionMap, RadiationTab radiationTab) {
        this.globalSelectionMap = globalSelectionMap;
        this.radiationTab = radiationTab;
    }

    public void addRoom(Room room) {
        rooms.add(room);
        fireTableRowsInserted(rooms.size()-1, rooms.size()-1);
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Room room = rooms.get(rowIndex);
        if (columnIndex == 0) {
            return globalSelectionMap.containsKey(room.getId())
                    ? globalSelectionMap.get(room.getId())
                    : room.isSelected(); // ← если в карте нет — используем сохранённое в модели
        } else {
            return room.getName();
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        if (columnIndex != 0) return;

        Room room = rooms.get(rowIndex);
        boolean val = (aValue instanceof Boolean) && (Boolean) aValue;

        globalSelectionMap.put(room.getId(), val); // для таблицы
        room.setSelected(val);                     // ПИШЕМ В МОДЕЛЬ (сохранится в проекте)

        fireTableRowsUpdated(rowIndex, rowIndex);
    }

    public void clear() {
        int size = rooms.size();
        rooms.clear();
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
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return columnIndex == 0;
    }
}