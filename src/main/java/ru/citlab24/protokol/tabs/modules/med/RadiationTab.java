package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
            "мусорокамера", "су", "с/у"
    );
    private static final List<String> KITCHEN_KEYWORDS = Arrays.asList(
            "кухня", "кухня-ниша", "кухня-гостиная", "кухня гостиная", "кухня ниша"
    );

    private Building radiationBuilding;
    private Map<Integer, Room> originalRoomMap = new HashMap<>(); // Связь ID оригинала -> комната
    private static final Logger logger = LoggerFactory.getLogger(RadiationTab.class);
    private JTable spaceTable;
    private SpaceTableModel spaceTableModel;
    private Building currentBuilding;
    private final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final Set<Integer> processedSpaces = new HashSet<>(); // Для отслеживания обработанных помещений

    private JList<Floor> floorList;
    private DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private JTable roomTable;
    private RadiationRoomsTableModel roomsTableModel;
    private JButton splitOfficeButton;
    private JButton splitRoomButton;

    public RadiationTab() {
        roomsTableModel = new RadiationRoomsTableModel(globalRoomSelectionMap);
        spaceTableModel = new SpaceTableModel();
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel mainPanel = new JPanel(new GridLayout(1, 3, 10, 10));
        mainPanel.add(createFloorPanel());
        mainPanel.add(createSpacePanel());
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

        // Создаем панель для списка и кнопки
        JPanel floorListPanel = new JPanel(new BorderLayout());
        floorListPanel.add(new JScrollPane(floorList), BorderLayout.CENTER);

        // Кнопка для обработки всех квартир
        JButton selectForAllButton = new JButton("Выбрать для всех квартир");
        selectForAllButton.addActionListener(e -> processAllResidentialSpacesOnCurrentFloor());

        // Панель для кнопки (для выравнивания)
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(selectForAllButton);

        floorListPanel.add(buttonPanel, BorderLayout.SOUTH);

        panel.add(floorListPanel, BorderLayout.CENTER);
        return panel;
    }

    private JPanel createSpacePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения на этаже", SPACE_PANEL_COLOR));

        // Таблица помещений
        spaceTable = new JTable(spaceTableModel);
        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateRoomList(); // Обновляем комнаты при выборе помещения
            }
        });
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }

    private void splitOfficeSpace(Space officeSpace, List<SplitOfficeDialog.RoomData> roomDataList) {
        Floor floor = findParentFloor(officeSpace);
        if (floor == null) return;

        // Удаляем исходное помещение
        floor.getSpaces().remove(officeSpace);

        // Создаем новые помещения
        for (SplitOfficeDialog.RoomData data : roomDataList) {
            Space newSpace = new Space();
            newSpace.setIdentifier(data.getSpaceIdentifier());
            newSpace.setType(Space.SpaceType.OFFICE);

            Room newRoom = new Room();
            newRoom.setName(data.getRoomName());
            newRoom.setId(generateUniqueRoomId());
            newSpace.addRoom(newRoom);

            floor.addSpace(newSpace);

            // Автоматически выбираем комнату
            globalRoomSelectionMap.put(newRoom.getId(), true);
        }

        // Обновляем UI
        refreshData();
        updateBuildingTab(floor);
    }

    private void updateBuildingTab(Floor updatedFloor) {
        Window mainFrame = SwingUtilities.getWindowAncestor(this);
        if (mainFrame instanceof MainFrame) {
            ((MainFrame) mainFrame).getBuildingTab().refreshFloor(updatedFloor);
        }
    }

    private JPanel createRoomPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Комнаты на этаже", ROOM_PANEL_COLOR));

        // Создаем кнопку "Добавить точки" для комнат
        splitRoomButton = new JButton("Добавить точки");
        splitRoomButton.setVisible(false);
        splitRoomButton.addActionListener(this::splitRoomAction);

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(splitRoomButton);

        // Таблица комнат с чекбоксами
        roomTable = new JTable(roomsTableModel);

        // Слушатель выбора комнаты
        roomTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                int row = roomTable.getSelectedRow();
                splitRoomButton.setVisible(row >= 0);
            }
        });

        panel.add(buttonPanel, BorderLayout.NORTH);
        panel.add(new JScrollPane(roomTable), BorderLayout.CENTER);
        return panel;
    }
    private void splitRoomAction(ActionEvent e) {
        int row = roomTable.getSelectedRow();
        if (row < 0) return;

        Room selectedRoom = roomsTableModel.getRoomAt(row);
        if (selectedRoom == null) return;

        // Запрос количества точек
        String input = JOptionPane.showInputDialog(
                this,
                "Сколько точек замеров будет в этом помещении?",
                "Количество точек",
                JOptionPane.QUESTION_MESSAGE
        );

        if (input == null) return; // пользователь отменил

        int pointCount;
        try {
            pointCount = Integer.parseInt(input);
            if (pointCount <= 0) {
                JOptionPane.showMessageDialog(this, "Введите положительное число", "Ошибка", JOptionPane.ERROR_MESSAGE);
                return;
            }
        } catch (NumberFormatException ex) {
            JOptionPane.showMessageDialog(this, "Введите целое число", "Ошибка", JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Диалог для ввода суффиксов
        SuffixInputDialog dialog = new SuffixInputDialog(
                (Frame) SwingUtilities.getWindowAncestor(this),
                selectedRoom.getName(),
                pointCount
        );
        dialog.setVisible(true);

        if (dialog.isConfirmed()) {
            splitRoom(selectedRoom, dialog.getSuffixes());
        }
    }
    private void splitRoom(Room selectedRoom, List<String> suffixes) {
        Space space = findParentSpace(selectedRoom);
        if (space == null) return;

        // Сохраняем состояние выбранности оригинала
        boolean wasSelected = globalRoomSelectionMap.getOrDefault(selectedRoom.getId(), false);

        space.getRooms().remove(selectedRoom);
        globalRoomSelectionMap.remove(selectedRoom.getId());

        for (String suffix : suffixes) {
            Room newRoom = new Room();
            newRoom.setId(generateUniqueRoomId());
            newRoom.setName(selectedRoom.getName() + suffix);
            newRoom.setOriginalRoomId(selectedRoom.getOriginalRoomId());

            // Сохраняем состояние оригинала для копии
            globalRoomSelectionMap.put(newRoom.getId(), wasSelected);
            space.addRoom(newRoom);
        }

        refreshData();
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
        }
    }

    private void updateRoomList() {
        roomsTableModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        int selectedSpaceRow = spaceTable.getSelectedRow();

        if (selectedFloor != null && selectedSpaceRow >= 0) {
            Space selectedSpace = spaceTableModel.getSpaceAt(selectedSpaceRow);

            for (Room room : selectedSpace.getRooms()) {
                roomsTableModel.addRoom(room);
            }

            // ВОССТАНАВЛИВАЕМ ПРОВЕРКУ НА ПЕРВОЕ ПОМЕЩЕНИЕ
            if (isResidentialSpace(selectedSpace) &&
                    !processedSpaces.contains(selectedSpace.getId()) &&
                    isFirstResidentialSpaceOnFloor(selectedSpace, selectedFloor)) {

                applyRoomSelectionRulesForResidentialSpace(selectedSpace);
            }
        }
        roomsTableModel.fireTableDataChanged();
    }
    private boolean isFirstResidentialSpaceOnFloor(Space space, Floor floor) {
        for (Space sp : floor.getSpaces()) {
            if (isResidentialSpace(sp)) {
                return sp.equals(space);
            }
        }
        return false;
    }

    private void processAllResidentialSpacesOnCurrentFloor() {
        Floor currentFloor = floorList.getSelectedValue();
        if (currentFloor == null) return;

        for (Space space : currentFloor.getSpaces()) {
            if (isResidentialSpace(space)) {
                // Убираем проверку на firstResidentialSpace
                applyRoomSelectionRulesForResidentialSpace(space);
                processedSpaces.add(space.getId());
            }
        }
        roomsTableModel.fireTableDataChanged();
    }

    private void applyRoomSelectionRulesForResidentialSpace(Space space) {
        if (processedSpaces.contains(space.getIdentifier())) return;
        try {
            // Получаем список всех комнат в помещении, не исключенных
            List<Room> validRooms = space.getRooms().stream()
                    .filter(room -> room != null && !containsAny(room.getName().toLowerCase(), EXCLUDED_ROOMS))
                    .collect(Collectors.toList());

            // Разделяем на кухни и не-кухни
            List<Room> kitchenRooms = validRooms.stream()
                    .filter(room -> containsAny(room.getName().toLowerCase(), KITCHEN_KEYWORDS))
                    .collect(Collectors.toList());

            List<Room> otherRooms = validRooms.stream()
                    .filter(room -> !containsAny(room.getName().toLowerCase(), KITCHEN_KEYWORDS))
                    .collect(Collectors.toList());

            // Сбрасываем все галочки в этом помещении
            for (Room room : space.getRooms()) {
                if (room != null) {
                    globalRoomSelectionMap.put(room.getId(), false);
                }
            }

            // Правило выбора двух комнат
            if (kitchenRooms.isEmpty()) {
                // Нет кухонь -> отмечаем две не-кухни (если есть)
                int count = 0;
                for (Room room : otherRooms) {
                    if (room != null) {
                        globalRoomSelectionMap.put(room.getId(), true);
                        count++;
                        if (count >= 2) break;
                    }
                }
            } else {
                // Отмечаем одну кухню (первую)
                if (kitchenRooms.get(0) != null) {
                    globalRoomSelectionMap.put(kitchenRooms.get(0).getId(), true);
                }

                // Отмечаем одну не-кухню (если есть) или вторую кухню (если не-кухонь нет)
                if (!otherRooms.isEmpty() && otherRooms.get(0) != null) {
                    globalRoomSelectionMap.put(otherRooms.get(0).getId(), true);
                } else if (kitchenRooms.size() > 1 && kitchenRooms.get(1) != null) {
                    globalRoomSelectionMap.put(kitchenRooms.get(1).getId(), true);
                }
            }

            // Помечаем пространство как обработанное
            processedSpaces.add(space.getId());

        } catch (Exception e) {
            logger.error("Ошибка при обработке помещения {}: {}", space.getIdentifier(), e.getMessage());
        }
    }


    public void setRoomSelectionState(int roomId, boolean state) {
        globalRoomSelectionMap.put(roomId, state);
        roomsTableModel.fireTableDataChanged();
    }


    private void handleFirstTwoFloors(Space space) {
        logger.info("Обработка помещения {} по правилам первых двух этажей", space.getIdentifier());

        List<Room> validRooms = space.getRooms().stream()
                .filter(room -> !containsAny(room.getName().toLowerCase(), EXCLUDED_ROOMS))
                .collect(Collectors.toList());

        List<Room> kitchenRooms = validRooms.stream()
                .filter(room -> containsAny(room.getName().toLowerCase(), KITCHEN_KEYWORDS))
                .collect(Collectors.toList());

        // Сбрасываем все галочки
        space.getRooms().forEach(room ->
                globalRoomSelectionMap.put(room.getId(), false));

        // Выбираем кухню
        if (!kitchenRooms.isEmpty()) {
            Room kitchen = kitchenRooms.get(0);
            globalRoomSelectionMap.put(kitchen.getId(), true);
            logger.info("Выбрана кухня: {}", kitchen.getName());
        }

        // Выбираем еще одну комнату (не кухню)
        Optional<Room> otherRoom = validRooms.stream()
                .filter(room -> !containsAny(room.getName().toLowerCase(), KITCHEN_KEYWORDS))
                .findFirst();

        if (otherRoom.isPresent()) {
            Room room = otherRoom.get();
            globalRoomSelectionMap.put(room.getId(), true);
            logger.info("Выбрана комната: {}", room.getName());
        }
    }

    public void copyRoomSelectionState(int originalRoomId, int newRoomId) {
        if (globalRoomSelectionMap.containsKey(originalRoomId)) {
            globalRoomSelectionMap.put(newRoomId, globalRoomSelectionMap.get(originalRoomId));
        }
    }
    // Новый метод для определения жилых помещений
    private boolean isResidentialSpace(Space space) {
        if (space == null) return false;
        return space.getType() == Space.SpaceType.APARTMENT;
    }
    public static boolean isExcludedRoom(String roomName) {
        String lowerName = roomName.toLowerCase();
        return EXCLUDED_ROOMS.stream().anyMatch(lowerName::contains);
    }

    private boolean containsAny(String source, List<String> targets) {
        for (String target : targets) {
            if (source.contains(target)) {
                return true;
            }
        }
        return false;
    }

    private void forceSelectOfficeRooms() {
        if (currentBuilding == null) return;

        for (Floor floor : currentBuilding.getFloors()) {
            for (Space space : floor.getSpaces()) {
                if (space.getType() == Space.SpaceType.OFFICE) {
                    for (Room room : space.getRooms()) {
                        if (!isExcludedRoom(room.getName())) {
                            globalRoomSelectionMap.put(room.getId(), true);
                        }
                    }
                }
            }
        }
    }

    public void setBuilding(Building building) {
        // Сохраняем текущие состояния перед обновлением
        Map<String, Boolean> savedSelections = saveSelections();

        this.currentBuilding = building;
        processedSpaces.clear();

        // Восстанавливаем состояния после обновления
        restoreSelections(savedSelections);
        forceSelectOfficeRooms();

        // Инициализация оригинальных комнат
        for (Floor floor : building.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    originalRoomMap.put(room.getId(), room);
                }
            }
        }

        radiationBuilding = cloneBuilding(building);
        refreshFloors();
    }

    private Building cloneBuilding(Building original) {
        Building clone = new Building();
        // ... копирование свойств ...
        for (Floor floor : original.getFloors()) {
            clone.addFloor(cloneFloor(floor));
        }
        return clone;
    }

    private Floor cloneFloor(Floor original) {
        Floor clone = new Floor();
        // ... копирование свойств ...
        for (Space space : original.getSpaces()) {
            clone.addSpace(cloneSpace(space));
        }
        return clone;
    }

    private Space cloneSpace(Space original) {
        Space clone = new Space();
        // ... копирование свойств ...
        for (Room room : original.getRooms()) {
            clone.addRoom(cloneRoom(room));
        }
        return clone;
    }

    private Room cloneRoom(Room original) {
        Room clone = new Room();
        clone.setId(generateUniqueRoomId());
        clone.setName(original.getName());
        clone.setOriginalRoomId(original.getId()); // Связь с оригиналом
        return clone;
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding != null) {
            currentBuilding.getFloors().forEach(floorListModel::addElement);
            if (!floorListModel.isEmpty()) {
                floorList.setSelectedIndex(0);
                // Принудительно обновляем помещения и комнаты
                SwingUtilities.invokeLater(() -> {
                    updateSpaceList();
                    updateRoomList();
                });
            }
        }
    }

    public void refreshData() {
        String selectedFloorNumber = null;
        String selectedSpaceIdentifier = null;

        if (floorList.getSelectedValue() != null) {
            selectedFloorNumber = floorList.getSelectedValue().getNumber();
        }

        // FIX: Use getSelectedSpace() helper
        Space selectedSpace = getSelectedSpace();
        if (selectedSpace != null) {
            // FIX: Use identifier (String) instead of ID (Integer)
            selectedSpaceIdentifier = selectedSpace.getIdentifier();
        }

        // Сохраняем состояния комнат
        Map<Integer, Boolean> savedSelection = new HashMap<>(globalRoomSelectionMap);
        processedSpaces.clear();

        // Обновляем список этажей
        refreshFloors();

        // Восстанавливаем выделение этажа
        restoreFloorSelection(selectedFloorNumber);

        // Обновляем таблицу помещений
        updateSpaceList();

        // Восстанавливаем выделение помещения
        restoreSpaceSelection(selectedSpaceIdentifier);

        // Обновляем таблицу комнат
        updateRoomList();

        // Восстанавливаем состояния комнат
        globalRoomSelectionMap.putAll(savedSelection);
        roomsTableModel.fireTableDataChanged();
    }
    private Space getSelectedSpace() {
        int row = spaceTable.getSelectedRow();
        if (row >= 0) {
            return spaceTableModel.getSpaceAt(row);
        }
        return null;
    }
    // Восстановление этажа по номеру
    private void restoreFloorSelection(String floorNumber) {
        if (floorNumber == null) return;

        for (int i = 0; i < floorListModel.size(); i++) {
            if (floorListModel.get(i).getNumber().equals(floorNumber)) {
                floorList.setSelectedIndex(i);
                return;
            }
        }
    }
    public void selectSpaceByIndex(int index) {
        if (index >= 0 && index < spaceTableModel.getRowCount()) {
            spaceTable.setRowSelectionInterval(index, index);
            updateRoomList();
        }
    }

    // Восстановление помещения по идентификатору
    private void restoreSpaceSelection(String spaceIdentifier) {
        if (spaceIdentifier == null) {
            if (spaceTableModel.getRowCount() > 0) {
                spaceTable.setRowSelectionInterval(0, 0);
            }
            return;
        }

        for (int i = 0; i < spaceTableModel.getRowCount(); i++) {
            if (spaceTableModel.getSpaceAt(i).getIdentifier().equals(spaceIdentifier)) {
                spaceTable.setRowSelectionInterval(i, i);
                return;
            }
        }

        if (spaceTableModel.getRowCount() > 0) {
            spaceTable.setRowSelectionInterval(0, 0);
        }
    }

    public Map<String, Boolean> saveSelections() {
        Map<String, Boolean> selections = new HashMap<>();
        for (Map.Entry<Integer, Boolean> entry : globalRoomSelectionMap.entrySet()) {
            Room room = findRoomById(entry.getKey());
            if (room != null) {
                Space space = findParentSpace(room);
                Floor floor = findParentFloor(space);
                if (floor != null && space != null) {
                    String key = floor.getNumber() + "|" + space.getIdentifier() + "|" + room.getName();
                    selections.put(key, entry.getValue());
                }
            }
        }
        return selections;
    }

    public void restoreSelections(Map<String, Boolean> savedSelections) {
        Map<Integer, Boolean> newSelectionMap = new HashMap<>();
        for (Room room : getAllRooms()) {
            String key = generateRoomKey(room);
            if (savedSelections.containsKey(key)) {
                newSelectionMap.put(room.getId(), savedSelections.get(key));
            } else {
                Boolean currentState = globalRoomSelectionMap.get(room.getId());
                newSelectionMap.put(room.getId(), currentState != null ? currentState : false);
            }
        }
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
        if (radiationBuilding == null) return null;
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
    public Map<Integer, Boolean> getFloorSelections(Floor floor) {
        Map<Integer, Boolean> selections = new HashMap<>();
        for (Space space : floor.getSpaces()) {
            for (Room room : space.getRooms()) {
                if (globalRoomSelectionMap.containsKey(room.getId())) {
                    selections.put(room.getId(), globalRoomSelectionMap.get(room.getId()));
                }
            }
        }
        return selections;
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
    private int generateUniqueRoomId() {
        return UUID.randomUUID().hashCode();
    }
}