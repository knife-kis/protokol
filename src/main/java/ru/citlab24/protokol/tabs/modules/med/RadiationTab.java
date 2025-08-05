package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.MainFrame;
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
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }

    private void splitOfficeAction(ActionEvent e) {
        int row = spaceTable.getSelectedRow();
        if (row < 0) return;

        Space selectedSpace = spaceTableModel.getSpaceAt(row);
        if (selectedSpace == null || selectedSpace.getType() != Space.SpaceType.OFFICE) return;

        // Диалог для разделения помещения
        SplitOfficeDialog dialog = new SplitOfficeDialog(
                (Frame) SwingUtilities.getWindowAncestor(this),
                selectedSpace.getIdentifier()
        );
        dialog.setVisible(true);

        if (dialog.isConfirmed()) {
            splitOfficeSpace(selectedSpace, dialog.getRoomDataList());
        }
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
        // ... существующая настройка таблицы ...

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
        Space space = findParentSpace(selectedRoom); // В radiationBuilding
        if (space == null) return;

        // Сохраняем состояние выбранности
        boolean wasSelected = globalRoomSelectionMap.getOrDefault(selectedRoom.getId(), false);

        // Удаляем клонированную комнату
        space.getRooms().remove(selectedRoom);
        globalRoomSelectionMap.remove(selectedRoom.getId());

        // Создаем новые клонированные комнаты
        for (String suffix : suffixes) {
            Room newRoom = new Room();
            newRoom.setId(generateUniqueRoomId());
            newRoom.setName(selectedRoom.getName() + suffix);
            newRoom.setOriginalRoomId(selectedRoom.getOriginalRoomId()); // Связь с оригиналом

            space.addRoom(newRoom);
            globalRoomSelectionMap.put(newRoom.getId(), wasSelected);
        }

        refreshData(); // Обновляем UI
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
                logger.info("Selected space: {} (type: {})",
                        selectedSpace.getIdentifier(), selectedSpace.getType());

                // Для ВСЕХ типов помещений инициализируем выбор комнат
                for (Room room : selectedSpace.getRooms()) {
                    int roomId = room.getId();
                    if (!globalRoomSelectionMap.containsKey(roomId)) {
                        initRoomSelection(room);
                    }
                    roomsTableModel.addRoom(room);
                }

                // Дополнительная обработка только для жилых помещений
                if (isResidentialSpace(selectedSpace)) {
                    logger.info("Processing as residential space");
                    if (!processedSpaces.contains(selectedSpace.getId())) {
                        processFirstResidentialSpace(selectedSpace);
                        processedSpaces.add(selectedSpace.getId());
                    }
                }
            }
        }
        roomsTableModel.fireTableDataChanged();
    }


    private void processFirstResidentialSpace(Space space) {
        if (!processedSpaces.contains(space.getId())) {
            applyRoomSelectionRulesForResidentialSpace(space);
            processedSpaces.add(space.getId());
        }
    }

    private void initRoomSelection(Room room) {
        int roomId = room.getId();
        String roomName = room.getName().toLowerCase();

        // Всегда отключаем исключенные комнаты
        if (isExcludedRoom(roomName)) {
            globalRoomSelectionMap.put(roomId, false);
            return;
        }

        Space space = findParentSpace(room);
        if (space == null) {
            globalRoomSelectionMap.put(roomId, false);
            return;
        }

        // Для офисных помещений - всегда включено
        if (space.getType() == Space.SpaceType.OFFICE) {
            globalRoomSelectionMap.put(roomId, true);
            return;
        }

        // Для общественных помещений - всегда отключено
        if (space.getType() == Space.SpaceType.PUBLIC_SPACE) {
            globalRoomSelectionMap.put(roomId, false);
            return;
        }

        // Для жилых помещений по умолчанию отключаем
        // (далее будет обработано в processFirstResidentialSpace)
        globalRoomSelectionMap.put(roomId, false);
    }

    private void processAllResidentialSpacesOnCurrentFloor() {
        Floor currentFloor = floorList.getSelectedValue();
        if (currentFloor == null) {
            logger.info("Не выбран этаж для обработки");
            return;
        }

        logger.info("Обработка всех квартир на этаже: {}", currentFloor.getNumber());

        boolean changesMade = false;
        for (Space space : currentFloor.getSpaces()) {
            if (space != null && isResidentialSpace(space)) {
                logger.info("Применение правил для квартиры: {}", space.getIdentifier());
                applyRoomSelectionRulesForResidentialSpace(space);
                changesMade = true;
            }
        }

        if (changesMade) {
            // Обновляем таблицу комнат
            roomsTableModel.fireTableDataChanged();
            // Обновляем список помещений
            updateSpaceList();
            // Пересчитываем представление
            revalidate();
            repaint();

            logger.info("Правила применены для всех квартир на этаже");
        } else {
            logger.info("На этаже нет квартир для обработки");
        }
    }

    private void applyRoomSelectionRulesForResidentialSpace(Space space) {
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

    private void processResidentialSpace(Space space, Floor floor) {
        // Проверяем, является ли это помещение первым жилым на этаже
        if (!isFirstResidentialSpaceOnFloor(space, floor)) {
            // Не первое помещение - все комнаты отключаем
            for (Room r : space.getRooms()) {
                globalRoomSelectionMap.put(r.getId(), false);
            }
            return;
        }

        // Первое жилое помещение на этаже
        List<Room> validRooms = space.getRooms().stream()
                .filter(r -> !containsAny(r.getName().toLowerCase(), EXCLUDED_ROOMS))
                .collect(Collectors.toList());

        // Находим кухню
        Optional<Room> kitchenOpt = validRooms.stream()
                .filter(r -> containsAny(r.getName().toLowerCase(), KITCHEN_KEYWORDS))
                .findFirst();

        // Находим другую комнату (не кухню)
        Optional<Room> otherRoomOpt = validRooms.stream()
                .filter(r -> !containsAny(r.getName().toLowerCase(), KITCHEN_KEYWORDS))
                .findFirst();

        // Устанавливаем все комнаты в false
        for (Room r : space.getRooms()) {
            globalRoomSelectionMap.put(r.getId(), false);
        }

        // Отмечаем кухню (если есть)
        kitchenOpt.ifPresent(kitchen ->
                globalRoomSelectionMap.put(kitchen.getId(), true));

        // Отмечаем другую комнату (если есть)
        otherRoomOpt.ifPresent(room ->
                globalRoomSelectionMap.put(room.getId(), true));
    }

    public void setRoomSelectionState(int roomId, boolean state) {
        globalRoomSelectionMap.put(roomId, state);
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
    private void handleUpperFloors(Floor floor, Space currentSpace) {
        // Находим первое жилое помещение на этаже
        Optional<Space> firstSpace = floor.getSpaces().stream()
                .filter(this::isResidentialSpace)
                .findFirst();

        // Если текущее помещение - первое на этаже
        if (firstSpace.isPresent() && firstSpace.get().equals(currentSpace)) {
            handleFirstTwoFloors(currentSpace); // Та же логика, что и для первых этажей
        } else {
            // Для всех остальных помещений на этаже снимаем галочки
            currentSpace.getRooms().forEach(room ->
                    globalRoomSelectionMap.put(room.getId(), false));
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

    public void setBuilding(Building building) {
        // Сохраняем текущие состояния
        Map<String, Boolean> savedSelections = saveSelections();

        this.currentBuilding = building;
        processedSpaces.clear();

        // Восстанавливаем состояния для существующих комнат
        restoreSelections(savedSelections);

        for (Floor floor : building.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    originalRoomMap.put(room.getId(), room);
                }
            }
        }

        // Глубокое клонирование здания
        radiationBuilding = cloneBuilding(building);
        refreshFloors();
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
            for (Floor floor : currentBuilding.getFloors()) {
                floorListModel.addElement(floor);
            }
            if (!floorListModel.isEmpty()) {
                floorList.setSelectedIndex(0);
            }
        }
    }

    public void refreshData() {
        Map<Integer, Boolean> savedSelection = new HashMap<>(globalRoomSelectionMap);
        processedSpaces.clear(); // Сбрасываем при обновлении

        refreshFloors();
        globalRoomSelectionMap.putAll(savedSelection);

        if (roomTable != null) {
            roomsTableModel.fireTableDataChanged();
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

    private int generateUniqueSpaceId() {
        return UUID.randomUUID().hashCode();
    }
}