package ru.citlab24.protokol.tabs.modules.microclimateTab;

import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.table.AbstractTableModel;
import java.util.*;
import java.util.function.Consumer;

/** Таблица комнат для «Микроклимата»: [чекбокс | комната | нар. стены 0..4] */
final class MicroclimateRoomsTableModel extends AbstractTableModel {

    private final String[] COLUMN_NAMES = {"Измерения", "Комната", "Нар. стены"};
    private final Class<?>[] COLUMN_TYPES = {Boolean.class, String.class, Integer.class};

    private final Map<Integer, Boolean> globalSelectionMap; // по id комнаты
    private final List<Room> rooms = new ArrayList<>();
    private final Consumer<Integer> onUserTouched; // зовём, когда юзер кликает чекбокс

    // Ключевые слова «влажных»/санитарных помещений — по умолчанию 0 внешних стен
    private static final List<String> WET_KEYWORDS = Arrays.asList(
            "санузел", "сан узел", "сан. узел", "туалет", "с/у", "су",
            "ванная", "душ", "душевая", "совмещенный", "совмещённый",
            "моечная", "уборная", "санитар", "wc"
    );

    MicroclimateRoomsTableModel(Map<Integer, Boolean> globalSelectionMap,
                                Consumer<Integer> onUserTouched) {
        this.globalSelectionMap = globalSelectionMap;
        this.onUserTouched = onUserTouched;
    }

    void setRooms(List<Room> newRooms) {
        rooms.clear();
        if (newRooms != null) rooms.addAll(newRooms);
        fireTableDataChanged();
    }

    List<Room> getRooms() { return rooms; }

    Room getRoomAt(int row) {
        return (row >= 0 && row < rooms.size()) ? rooms.get(row) : null;
    }

    @Override public int getRowCount() { return rooms.size(); }
    @Override public int getColumnCount() { return COLUMN_NAMES.length; }
    @Override public String getColumnName(int column) { return COLUMN_NAMES[column]; }
    @Override public Class<?> getColumnClass(int columnIndex) { return COLUMN_TYPES[columnIndex]; }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        // чекбокс и «Нар. стены» редактируемы
        return columnIndex == 0 || columnIndex == 2;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        Room room = rooms.get(rowIndex);
        switch (columnIndex) {
            case 0: // чекбокс
                return Boolean.valueOf(globalSelectionMap.getOrDefault(room.getId(), room.isSelected()));
            case 1: // имя
                return room.getName();
            case 2: { // нар. стены
                Integer v = room.getExternalWallsCount();
                if (v == null) {
                    v = isWetRoom(room.getName()) ? 0 : 1; // по умолчанию: санузлы/ванные = 0, остальные = 1
                    room.setExternalWallsCount(v);
                }
                return v;
            }
            default:
                return null;
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        Room room = rooms.get(rowIndex);
        switch (columnIndex) {
            case 0: {
                boolean val = (aValue instanceof Boolean) && (Boolean) aValue;
                globalSelectionMap.put(room.getId(), val);
                room.setSelected(val);
                if (onUserTouched != null) onUserTouched.accept(room.getId());
                fireTableRowsUpdated(rowIndex, rowIndex);
                break;
            }
            case 2: {
                Integer v = null;
                if (aValue instanceof Number) v = ((Number) aValue).intValue();
                if (v == null && aValue instanceof String) {
                    try { v = Integer.parseInt(((String) aValue).trim()); } catch (Exception ignore) {}
                }
                if (v == null) return;
                if (v < 0) v = 0;
                if (v > 4) v = 4;
                room.setExternalWallsCount(v);
                fireTableRowsUpdated(rowIndex, rowIndex);
                break;
            }
        }
    }

    private static boolean isWetRoom(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(Locale.ROOT);
        for (String kw : WET_KEYWORDS) {
            if (n.contains(kw)) return true;
        }
        return false;
    }
}
