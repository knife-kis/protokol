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

    private boolean autoApplyRulesOnDisplay = true;
    private static final Logger logger = LoggerFactory.getLogger(RadiationTab.class);
    private JTable spaceTable;
    private SpaceTableModel spaceTableModel;
    private Building currentBuilding;
    public final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final Set<Integer> processedSpaces = new HashSet<>(); // Для отслеживания обработанных помещений

    private JList<Floor> floorList;
    private DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private JTable roomTable;
    private RadiationRoomsTableModel roomsTableModel;
    private JButton splitRoomButton;

    public RadiationTab() {
        roomsTableModel = new RadiationRoomsTableModel(globalRoomSelectionMap, this);
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
    public Map<Integer, Boolean> globalRoomSelectionMap() {
        return globalRoomSelectionMap;
    }


    private void updateRoomList() {
        roomsTableModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        int selectedSpaceRow = spaceTable.getSelectedRow();

        if (selectedFloor != null && selectedSpaceRow >= 0) {
            Space selectedSpace = spaceTableModel.getSpaceAt(selectedSpaceRow);
            boolean isOffice = selectedSpace.getType() == Space.SpaceType.OFFICE;
            boolean isApartment = selectedSpace.getType() == Space.SpaceType.APARTMENT;

            // Сохраняем текущие состояния перед обновлением
            Map<Integer, Boolean> savedStates = new HashMap<>();
            for (Room room : selectedSpace.getRooms()) {
                savedStates.put(room.getId(), globalRoomSelectionMap.get(room.getId()));
            }

            // Заполняем таблицу комнат
            for (Room room : selectedSpace.getRooms()) {
                roomsTableModel.addRoom(room);
            }

            // Восстанавливаем сохраненные состояния
            for (Room room : selectedSpace.getRooms()) {
                Boolean savedState = savedStates.get(room.getId());
                if (savedState != null) {
                    globalRoomSelectionMap.put(room.getId(), savedState);
                }
            }

            // Применяем правила ТОЛЬКО для необработанных помещений
            if (autoApplyRulesOnDisplay && !processedSpaces.contains(selectedSpace.getId())) {
                if (isOffice) {
                    selectAllOfficeRooms(selectedSpace);
                } else if (isApartment && isFirstResidentialSpaceOnFloor(selectedSpace, selectedFloor)) {
                    applyRoomSelectionRulesForResidentialSpace(selectedSpace);
                }
            }
        }
        roomsTableModel.fireTableDataChanged();
    }
    private void selectAllOfficeRooms(Space space) {
        for (Room room : space.getRooms()) {
            if (!isExcludedRoom(room.getName())) {
                globalRoomSelectionMap.put(room.getId(), true);
            } else {
                globalRoomSelectionMap.put(room.getId(), false);
            }
        }
        processedSpaces.add(space.getId());
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
        if (processedSpaces.contains(space.getId())) return;
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

    public void forceSelectOfficeRooms() {
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

    public void setBuilding(Building building, boolean forceOfficeSelection) {
        // поведение по умолчанию оставляем прежним: авто-правила разрешены
        setBuilding(building, forceOfficeSelection, true);
    }
    public void setBuilding(Building building, boolean forceOfficeSelection, boolean autoApplyRules) {
        logger.info("RadiationTab.setBuilding() - Начало установки здания");
        this.currentBuilding = building;
        this.autoApplyRulesOnDisplay = autoApplyRules; // ← управляем автоприменением правил
        processedSpaces.clear();

        // Перенос состояний из модели
        globalRoomSelectionMap.clear();
        for (Floor floor : building.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    globalRoomSelectionMap.put(room.getId(), room.isSelected());
                }
            }
        }

        if (forceOfficeSelection) {
            forceSelectOfficeRooms();
        }

        refreshFloors();
        logger.info("Здание установлено. Комнат: {}", globalRoomSelectionMap.size());
    }

    public void updateRoomSelectionStates() {
        logger.info("RadiationTab.updateRoomSelectionStates() - Начало обновления состояний комнат");
        int updatedCount = 0;

        for (Floor floor : currentBuilding.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    // Получаем состояние чекбокса из глобальной карты
                    Boolean selected = globalRoomSelectionMap.get(room.getId());
                    if (selected != null) {
                        // Обновляем состояние комнаты
                        room.setSelected(selected);
                        logger.trace("Обновление комнаты {}: {} -> {}",
                                room.getName(), room.isSelected(), selected);
                        updatedCount++;
                    }
                }
            }
        }
        logger.info("RadiationTab.updateRoomSelectionStates() - Обновлено {} комнат", updatedCount);
    }

    public void refreshFloors() {
        String selectedFloorNumber = null;
        if (floorList.getSelectedValue() != null) {
            selectedFloorNumber = floorList.getSelectedValue().getNumber();
        }

        floorListModel.clear();
        if (currentBuilding != null) {
            currentBuilding.getFloors().forEach(floorListModel::addElement);
        }

        // Восстанавливаем выделение этажа
        if (selectedFloorNumber != null) {
            for (int i = 0; i < floorListModel.size(); i++) {
                if (floorListModel.get(i).getNumber().equals(selectedFloorNumber)) {
                    floorList.setSelectedIndex(i);
                    break;
                }
            }
        }

        // Обновляем таблицы помещений и комнат
        updateSpaceList();
        updateRoomList();
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
        logger.info("RadiationTab.saveSelections() - Сохранение состояний");
        Map<String, Boolean> selections = new HashMap<>();
        int count = 0;

        for (Map.Entry<Integer, Boolean> entry : globalRoomSelectionMap.entrySet()) {
            Room room = findRoomById(entry.getKey());
            if (room != null) {
                Space space = findParentSpace(room);
                Floor floor = findParentFloor(space);
                if (floor != null && space != null) {
                    String key = floor.getNumber() + "|" + space.getIdentifier() + "|" + room.getName();
                    selections.put(key, entry.getValue());
                    count++;
                    logger.trace("Сохранено состояние: {} = {}", key, entry.getValue());
                }
            }
        }
        logger.info("RadiationTab.saveSelections() - Сохранено {} состояний", count);
        return selections;
    }

    public void restoreSelections(Map<String, Boolean> savedSelections) {
        logger.info("RadiationTab.restoreSelections() - Восстановление состояний");
        logger.debug("Получено {} состояний для восстановления", savedSelections.size());

        Map<Integer, Boolean> newSelectionMap = new HashMap<>();
        int restoredCount = 0;
        int notFoundCount = 0;

        for (Room room : getAllRooms()) {
            String key = generateRoomKey(room);
            if (savedSelections.containsKey(key)) {
                newSelectionMap.put(room.getId(), savedSelections.get(key));
                restoredCount++;
                logger.trace("Восстановлено состояние для {}: {}", key, savedSelections.get(key));
            } else {
                Boolean currentState = globalRoomSelectionMap.get(room.getId());
                newSelectionMap.put(room.getId(), currentState != null ? currentState : false);
                notFoundCount++;
                logger.trace("Состояние не найдено для {}, используется текущее: {}", key, currentState);
            }
        }

        globalRoomSelectionMap.clear();
        globalRoomSelectionMap.putAll(newSelectionMap);
        roomsTableModel.fireTableDataChanged();

        logger.info("RadiationTab.restoreSelections() - Восстановлено: {}, Не найдено: {}", restoredCount, notFoundCount);
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
    public Space findParentSpace(Room room) {
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
    private int generateUniqueRoomId() {
        return UUID.randomUUID().hashCode();
    }
}