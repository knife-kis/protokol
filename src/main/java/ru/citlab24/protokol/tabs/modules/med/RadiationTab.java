package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.*;
import java.util.List;

public class RadiationTab extends JPanel {
    // Константы для цветов и стилей
    private static final Color FLOOR_PANEL_COLOR = new Color(0, 115, 200);
    private static final Color SPACE_PANEL_COLOR = new Color(76, 175, 80);
    private static final Color ROOM_PANEL_COLOR = new Color(156, 39, 176);
    private static final Font HEADER_FONT = new Font("Segoe UI", Font.BOLD, 14);
    private static final List<String> EXCLUDED_ROOMS = Arrays.asList(
            "санузел", "сан узел", "сан. узел",
            "ванная", "ванная комната",
            "совмещенный", "совмещенный санузел", "туалет",
            "су", "с/у"
    );
    private static final List<String> KITCHEN_KEYWORDS = Arrays.asList(
            "кухня", "кухня-ниша", "кухня-гостиная", "кухня гостиная", "кухня ниша"
    );

    private JTable spaceTable;
    private SpaceTableModel spaceTableModel;
    private Building currentBuilding;
    private final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();

    private JList<Floor> floorList;
    private DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private JTable roomTable;
    private RadiationRoomsTableModel roomsTableModel;

    public RadiationTab() {
        roomsTableModel = new RadiationRoomsTableModel(globalRoomSelectionMap);
        spaceTableModel = new SpaceTableModel();
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // Создаем основную панель с тремя колонками
        JPanel mainPanel = new JPanel(new GridLayout(1, 3, 10, 10));

        // Колонка 1: Здания и этажи
        mainPanel.add(createFloorPanel());

        // Колонка 2: Помещения на этаже
        mainPanel.add(createSpacePanel());

        // Колонка 3: Комнаты с чекбоксами
        mainPanel.add(createRoomPanel());

        add(mainPanel, BorderLayout.CENTER);

        SwingUtilities.invokeLater(() -> {
            if (currentBuilding != null) {
                refreshFloors();
            }
        });
    }

    private JPanel createFloorPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Этажи здания", FLOOR_PANEL_COLOR));

        // Список этажей
        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new FloorListRenderer());
        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateSpaceList();
                updateRoomList();
            }
        });

        panel.add(new JScrollPane(floorList), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createSpacePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения на этаже", SPACE_PANEL_COLOR));

        // Заменяем список на таблицу
        spaceTable = new JTable(spaceTableModel);

        // Настройка внешнего вида таблицы
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateRoomList();
            }
        });

        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createRoomPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Комнаты на этаже", ROOM_PANEL_COLOR));

        // Таблица комнат с чекбоксами
        roomTable = new JTable(roomsTableModel);

        // Настройка колонки с чекбоксами
        roomTable.getColumnModel().getColumn(0).setCellRenderer(roomTable.getDefaultRenderer(Boolean.class));
        roomTable.getColumnModel().getColumn(0).setCellEditor(roomTable.getDefaultEditor(Boolean.class));
        roomTable.getColumnModel().getColumn(0).setPreferredWidth(30);

        // Настройка колонки с названиями комнат
        DefaultTableCellRenderer renderer = new DefaultTableCellRenderer();
        renderer.setHorizontalAlignment(SwingConstants.LEFT);
        roomTable.getColumnModel().getColumn(1).setCellRenderer(renderer);

        panel.add(new JScrollPane(roomTable), BorderLayout.CENTER);
        return panel;
    }

    private TitledBorder createTitledBorder(String title, Color color) {
        return BorderFactory.createTitledBorder(
                null, title, TitledBorder.LEFT, TitledBorder.TOP, HEADER_FONT, color);
    }

    private void updateSpaceList() {
        spaceTableModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();

        if (selectedFloor != null) {
            for (Space space : selectedFloor.getSpaces()) {
                spaceTableModel.addSpace(space);
            }

            // Автовыбор первого элемента
            if (spaceTableModel.getRowCount() > 0) {
                spaceTable.setRowSelectionInterval(0, 0);
            }
        }
    }

    private void updateRoomList() {
        roomsTableModel.clear();
        int selectedRow = spaceTable.getSelectedRow();

        if (selectedRow >= 0) {
            Space selectedSpace = spaceTableModel.getSpaceAt(selectedRow);
            if (selectedSpace != null) {
                for (Room room : selectedSpace.getRooms()) {
                    int roomId = room.getId();

                    // Инициализация только если состояние не задано
                    if (!globalRoomSelectionMap.containsKey(roomId)) {
                        initRoomSelection(room);
                    }

                    roomsTableModel.addRoom(room);
                }
            }
        }
        roomsTableModel.fireTableDataChanged();
    }

    private void initRoomSelection(Room room) {
        int roomId = room.getId();
        String roomName = room.getName().toLowerCase();

        // 1. Always disable excluded rooms
        if (containsAny(roomName, EXCLUDED_ROOMS)) {
            globalRoomSelectionMap.put(roomId, false);
            return;
        }

        // 2. For residential floors apply special logic
        Space space = findParentSpace(room); // First find parent space
        Floor floor = null;
        if (space != null) {
            floor = findParentFloor(space); // Then find parent floor
        }

        if (floor != null && floor.getType() == Floor.FloorType.RESIDENTIAL) {
            initResidentialRoomSelection(room, floor);
        } else {
            // 3. For non-residential floors - default true
            globalRoomSelectionMap.put(roomId, true);
        }
    }

    private void initResidentialRoomSelection(Room currentRoom, Floor floor) {
        Space space = findParentSpace(currentRoom);
        if (space == null) return;

        int roomId = currentRoom.getId();
        String currentRoomName = currentRoom.getName().toLowerCase();

        // Соберем все комнаты в помещении
        List<Room> allRooms = space.getRooms();
        List<Room> validRooms = new ArrayList<>();
        List<Room> kitchenRooms = new ArrayList<>();

        // Фильтруем комнаты
        for (Room room : allRooms) {
            String name = room.getName().toLowerCase();
            if (containsAny(name, EXCLUDED_ROOMS)) {
                // Для исключенных комнат сразу ставим false
                globalRoomSelectionMap.put(room.getId(), false);
            } else {
                validRooms.add(room);
                if (containsAny(name, KITCHEN_KEYWORDS)) {
                    kitchenRooms.add(room);
                }
            }
        }

        // Выбираем комнаты для включения
        Set<Integer> selectedRooms = new HashSet<>();

        // 1. Выбираем первую кухню (если есть)
        if (!kitchenRooms.isEmpty()) {
            Room kitchen = kitchenRooms.get(0);
            selectedRooms.add(kitchen.getId());
            globalRoomSelectionMap.put(kitchen.getId(), true);
        }

        // 2. Выбираем первую НЕ кухню (если есть)
        for (Room room : validRooms) {
            if (!selectedRooms.contains(room.getId()) &&
                    !containsAny(room.getName().toLowerCase(), KITCHEN_KEYWORDS)) {
                selectedRooms.add(room.getId());
                globalRoomSelectionMap.put(room.getId(), true);
                break; // Только одну комнату
            }
        }

        // 3. Для текущей комнаты, если она не выбрана - ставим false
        if (!selectedRooms.contains(roomId)) {
            globalRoomSelectionMap.put(roomId, false);
        }
    }

    private boolean containsAny(String source, List<String> targets) {
        for (String target : targets) {
            if (source.contains(target)) {
                return true;
            }
        }
        return false;
    }

    public void setBuilding(Building building) {
        this.currentBuilding = building;
        // Не очищаем globalRoomSelectionMap, чтобы сохранить состояния между зданиями
        refreshFloors();;
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding != null) {
            for (Floor floor : currentBuilding.getFloors()) {
                floorListModel.addElement(floor);
            }
            if (!floorListModel.isEmpty()) {
                floorList.setSelectedIndex(0);
            }
        }
    }

    public void refreshData() {
        // Сохраняем текущее состояние выбранных комнат
        Map<Integer, Boolean> savedSelection = new HashMap<>(globalRoomSelectionMap);

        refreshFloors();

        // Восстанавливаем состояние после обновления
        // Убираем очистку globalRoomSelectionMap
        globalRoomSelectionMap.putAll(savedSelection);

        // Принудительно обновляем таблицу комнат
        if (roomTable != null) {
            roomsTableModel.fireTableDataChanged();
        }
    }

    public Map<String, Boolean> saveSelections() {
        Map<String, Boolean> selections = new HashMap<>();
        for (Map.Entry<Integer, Boolean> entry : globalRoomSelectionMap.entrySet()) {
            Room room = findRoomById(entry.getKey());
            if (room != null) {
                String key = generateRoomKey(room);
                selections.put(key, entry.getValue());
            }
        }
        return selections;
    }

    public void restoreSelections(Map<String, Boolean> savedSelections) {
        // Временная карта для новых состояний
        Map<Integer, Boolean> newSelectionMap = new HashMap<>();

        for (Room room : getAllRooms()) {
            String key = generateRoomKey(room);
            if (savedSelections.containsKey(key)) {
                newSelectionMap.put(room.getId(), savedSelections.get(key));
            } else {
                // Сохраняем текущее состояние, если оно есть
                Boolean currentState = globalRoomSelectionMap.get(room.getId());
                newSelectionMap.put(room.getId(), currentState != null ? currentState : true);
            }
        }

        // Заменяем глобальную карту
        globalRoomSelectionMap.clear();
        globalRoomSelectionMap.putAll(newSelectionMap);
        roomsTableModel.fireTableDataChanged();
    }

    // Генерация уникального ключа для комнаты
    private String generateRoomKey(Room room) {
        Space space = findParentSpace(room);
        if (space == null) return "unknown";

        Floor floor = findParentFloor(space);
        if (floor == null) return "unknown";

        return String.format("%s|%s|%s",
                floor.getNumber(),
                space.getIdentifier(),
                room.getName()
        );
    }

    // Вспомогательные методы для поиска
    private Space findParentSpace(Room room) {
        if (currentBuilding == null) return null;

        for (Floor floor : currentBuilding.getFloors()) {
            for (Space space : floor.getSpaces()) {
                if (space.getRooms().contains(room)) {
                    return space;
                }
            }
        }
        return null;
    }

    private Floor findParentFloor(Space space) {
        if (currentBuilding == null) return null;

        for (Floor floor : currentBuilding.getFloors()) {
            if (floor.getSpaces().contains(space)) {
                return floor;
            }
        }
        return null;
    }

    private List<Room> getAllRooms() {
        List<Room> allRooms = new ArrayList<>();
        if (currentBuilding == null) return allRooms;

        for (Floor floor : currentBuilding.getFloors()) {
            for (Space space : floor.getSpaces()) {
                allRooms.addAll(space.getRooms());
            }
        }
        return allRooms;
    }

    // Добавленный метод для поиска комнаты по ID
    private Room findRoomById(Integer roomId) {
        if (roomId == null || currentBuilding == null) return null;

        for (Room room : getAllRooms()) {
            if (roomId.equals(room.getId())) {
                return room;
            }
        }
        return null;
    }


    private static class SpaceTableModel extends AbstractTableModel {
        private final String[] COLUMN_NAMES = {"Название"};
        private final List<Space> spaces = new ArrayList<>();

        public void addSpace(Space space) {
            spaces.add(space);
            fireTableRowsInserted(spaces.size() - 1, spaces.size() - 1);
        }

        public void clear() {
            int size = spaces.size();
            spaces.clear();
            if (size > 0) fireTableRowsDeleted(0, size - 1);
        }

        public Space getSpaceAt(int rowIndex) {
            if (rowIndex >= 0 && rowIndex < spaces.size()) {
                return spaces.get(rowIndex);
            }
            return null;
        }

        @Override
        public int getRowCount() {
            return spaces.size();
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
        public Object getValueAt(int rowIndex, int columnIndex) {
            Space space = spaces.get(rowIndex);
            return space.getIdentifier();
        }
    }
}