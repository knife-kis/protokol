package ru.citlab24.protokol.tabs.modules.med;

import ru.citlab24.protokol.tabs.models.Section;
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
    // ===== Подсветка (радиация) ===============================================
    private static final java.awt.Color HL_GREEN = new java.awt.Color(232, 245, 233); // лёгкий зелёный
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
    private JList<Section> sectionList;
    private DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private JTable roomTable;
    private RadiationRoomsTableModel roomsTableModel;
    private JButton splitRoomButton;

    public RadiationTab() {
        roomsTableModel = new RadiationRoomsTableModel(globalRoomSelectionMap, this);
        spaceTableModel = new SpaceTableModel();
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel mainPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridy = 0;
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weighty = 1.0;

// Секции — узкая колонка (вес 1)
        gbc.gridx = 0;
        gbc.weightx = 1.0;                 // ← в 3 раза меньше остальных
        gbc.insets = new Insets(0, 0, 0, 10);
        mainPanel.add(createSectionPanel(), gbc);

// Этажи — обычная ширина (вес 3)
        gbc.gridx = 1;
        gbc.weightx = 3.0;
        mainPanel.add(createFloorPanel(), gbc);

// Помещения — обычная ширина (вес 3)
        gbc.gridx = 2;
        gbc.weightx = 3.0;
        mainPanel.add(createSpacePanel(), gbc);

// Комнаты — обычная ширина (вес 3)
        gbc.gridx = 3;
        gbc.weightx = 3.0;
        gbc.insets = new Insets(0, 0, 0, 0); // без правого отступа у последнего
        mainPanel.add(createRoomPanel(), gbc);

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

        // БЫЛО:
        // floorList.setCellRenderer(new FloorListRenderer());

        // СТАЛО: рендерер с подсветкой этажей, где есть хотя бы одна галочка радиации
        floorList.setCellRenderer(new ListCellRenderer<Floor>() {
            private final FloorListRenderer base = new FloorListRenderer();
            @Override
            public Component getListCellRendererComponent(JList<? extends Floor> list,
                                                          Floor value, int index,
                                                          boolean isSelected, boolean cellHasFocus) {
                Component c = base.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (!isSelected && c instanceof JComponent) {
                    boolean highlight = rad_hasAnyOnFloor(value);
                    ((JComponent) c).setOpaque(true);
                    c.setBackground(highlight ? HL_GREEN : UIManager.getColor("List.background"));
                }
                return c;
            }
        });

        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateSpaceList();
                updateRoomList();
            }
        });

        // Панель списка и кнопок
        JPanel floorListPanel = new JPanel(new BorderLayout());
        floorListPanel.add(new JScrollPane(floorList), BorderLayout.CENTER);

        JButton selectForAllButton = new JButton("Проставить чекбоксы на этаже");
        selectForAllButton.addActionListener(e -> applyRulesForWholeFloor());

        JButton exportBtn = new JButton("Экспорт в Excel");
        exportBtn.addActionListener(e -> exportToExcel());

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(selectForAllButton);
        buttonPanel.add(exportBtn);

        floorListPanel.add(buttonPanel, BorderLayout.SOUTH);

        panel.add(floorListPanel, BorderLayout.CENTER);
        return panel;
    }

    private void exportToExcel() {
        if (currentBuilding == null) {
            JOptionPane.showMessageDialog(this, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        // синхронизируем глобальную карту чекбоксов в room.isSelected()
        updateRoomSelectionStates();

        // -1 = экспорт всех блок-секций
        RadiationExcelExporter.export(currentBuilding, -1, this);
    }

    // NEW — колонка секций слева от этажей
    private JPanel createSectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Секции", FLOOR_PANEL_COLOR));

        sectionList = new JList<>(sectionListModel);
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setFixedCellHeight(28);

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                // при выборе секции — перерисовываем этажи/помещения/комнаты
                refreshFloors();
            }
        });

        panel.add(new JScrollPane(sectionList), BorderLayout.CENTER);
        return panel;
    }

    // NEW — подгрузка секций из currentBuilding в список
    private void refreshSections() {
        sectionListModel.clear();
        if (currentBuilding != null && currentBuilding.getSections() != null) {
            for (Section s : currentBuilding.getSections()) {
                sectionListModel.addElement(s);
            }
        }
        if (!sectionListModel.isEmpty()) {
            // если раньше было выделение — можно попытаться его сохранить,
            // но для простоты выберем первую секцию
            sectionList.setSelectedIndex(0);
        }
    }


    private JPanel createSpacePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения на этаже", SPACE_PANEL_COLOR));

        // Таблица помещений
        spaceTable = new JTable(spaceTableModel);
        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateRoomList();
            }
        });
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // Подсветка помещений, где есть хотя бы одна галочка радиации
        spaceTable.getColumnModel().getColumn(0).setCellRenderer(new javax.swing.table.DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value,
                                                           boolean isSelected, boolean hasFocus,
                                                           int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (!isSelected) {
                    Space s = spaceTableModel.getSpaceAt(row);
                    boolean highlight = rad_hasAnyInSpace(s);
                    c.setBackground(highlight ? HL_GREEN : UIManager.getColor("Table.background"));
                }
                return c;
            }
        });

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

            if (autoApplyRulesOnDisplay) {
                // Офисные этажи: проставляем во всех офисах (как раньше)
                // Офисные и смешанные этажи: проставляем во всех офисных помещениях
                if (selectedFloor.getType() == Floor.FloorType.OFFICE
                        || selectedFloor.getType() == Floor.FloorType.MIXED) {
                    for (Space space : selectedFloor.getSpaces()) {
                        if (space.getType() == Space.SpaceType.OFFICE && !processedSpaces.contains(space.getId())) {
                            selectAllOfficeRooms(space);
                        }
                    }
                }


                // Жилые и смешанные этажи: применяем правила к ПЕРВОЙ квартире, у которой уже есть комнаты
                Space firstApartmentWithRooms = selectedFloor.getSpaces().stream()
                        .filter(s -> s.getType() == Space.SpaceType.APARTMENT)
                        .filter(s -> !s.getRooms().isEmpty())
                        .findFirst()
                        .orElse(null);

                if (firstApartmentWithRooms != null && !processedSpaces.contains(firstApartmentWithRooms.getId())) {
                    applyRoomSelectionRulesForResidentialSpace(firstApartmentWithRooms);
                }
            }
        }
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
        if (space.getRooms().isEmpty()) {
            return; // не помечаем как обработанное, если нет комнат
        }
        for (Room room : space.getRooms()) {
            if (!isExcludedRoom(room.getName())) {
                globalRoomSelectionMap.put(room.getId(), true);
            } else {
                globalRoomSelectionMap.put(room.getId(), false);
            }
        }
        processedSpaces.add(space.getId());
        rad_refreshHighlights();
    }

    private boolean isFirstResidentialSpaceOnFloor(Space space, Floor floor) {
        for (Space sp : floor.getSpaces()) {
            if (isResidentialSpace(sp)) {
                return sp.equals(space);
            }
        }
        return false;
    }
    private void applyRulesForWholeFloor() {
        Floor currentFloor = floorList.getSelectedValue();
        if (currentFloor == null) return;

        for (Space space : currentFloor.getSpaces()) {
            if (space.getType() == Space.SpaceType.APARTMENT) {
                applyRoomSelectionRulesForResidentialSpace(space);
                processedSpaces.add(space.getId());
            } else if (space.getType() == Space.SpaceType.OFFICE) {
                selectAllOfficeRooms(space);
            } // PUBLIC — ничего не трогаем
        }
        roomsTableModel.fireTableDataChanged();
        rad_refreshHighlights();
    }

    private void applyRoomSelectionRulesForResidentialSpace(Space space) {
        if (processedSpaces.contains(space.getId())) return;
        try {
            List<Room> validRooms = space.getRooms().stream()
                    .filter(Objects::nonNull)
                    .filter(r -> !containsAny(r.getName().toLowerCase(), EXCLUDED_ROOMS))
                    .collect(Collectors.toList());

            if (validRooms.isEmpty()) {
                return; // не помечаем как обработанное — ждём комнат
            }

            // Сбросим все галочки в помещении
            for (Room room : space.getRooms()) {
                if (room != null) globalRoomSelectionMap.put(room.getId(), false);
            }

            List<Room> kitchenRooms = validRooms.stream()
                    .filter(r -> containsAny(r.getName().toLowerCase(), KITCHEN_KEYWORDS))
                    .collect(Collectors.toList());
            List<Room> otherRooms = validRooms.stream()
                    .filter(r -> !containsAny(r.getName().toLowerCase(), KITCHEN_KEYWORDS))
                    .collect(Collectors.toList());

            if (kitchenRooms.isEmpty()) {
                int count = 0;
                for (Room room : otherRooms) {
                    globalRoomSelectionMap.put(room.getId(), true);
                    if (++count >= 2) break;
                }
            } else {
                globalRoomSelectionMap.put(kitchenRooms.get(0).getId(), true);
                if (!otherRooms.isEmpty()) {
                    globalRoomSelectionMap.put(otherRooms.get(0).getId(), true);
                } else if (kitchenRooms.size() > 1) {
                    globalRoomSelectionMap.put(kitchenRooms.get(1).getId(), true);
                }
            }

            processedSpaces.add(space.getId());
            rad_refreshHighlights();
        } catch (Exception e) {
            logger.error("Ошибка при обработке помещения {}: {}", space.getIdentifier(), e.getMessage());
        }
    }

    private boolean hasResidentialSelectionOnFloor(Floor floor) {
        // Есть ли уже хоть одна отмеченная не-санузловая комната в любой квартире этого этажа?
        for (Space s : floor.getSpaces()) {
            if (s.getType() != Space.SpaceType.APARTMENT) continue;
            for (Room r : s.getRooms()) {
                // сперва смотрим в живую карту чекбоксов, затем в сохранённое состояние комнаты
                Boolean v = globalRoomSelectionMap.get(r.getId());
                boolean chosen = (v != null) ? v : r.isRadiationSelected(); // ← fallback к сохранённому флагу радиации
                if (chosen) {
                    String n = safe(r.getName());
                    if (!isExcludedRoom(n)) return true;
                }

            }
        }
        return false;
    }


    private static String safe(String s) {
        return s == null ? "" : s.trim().toLowerCase(Locale.ROOT);
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
                        } else {
                            globalRoomSelectionMap.put(room.getId(), false);
                        }
                    }
                    processedSpaces.add(space.getId()); // ← помечаем как обработанное
                }
            }
        }
    }

    public void setBuilding(Building building, boolean forceOfficeSelection, boolean autoApplyRules) {
        logger.info("RadiationTab.setBuilding() - Начало установки здания");
        this.currentBuilding = building;
        this.autoApplyRulesOnDisplay = autoApplyRules; // ← управляем автоприменением правил
        processedSpaces.clear();

        // Перенос состояний из модели
        // Перенос состояний из модели (РАДИАЦИЯ — своё поле)
        globalRoomSelectionMap.clear();
        for (Floor floor : building.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    globalRoomSelectionMap.put(room.getId(), room.isRadiationSelected());
                }
            }
        }
        if (forceOfficeSelection) {
            forceSelectOfficeRooms();
        }

        refreshSections();
        refreshFloors();
        logger.info("Здание установлено. Комнат: {}", globalRoomSelectionMap.size());
    }

    public void updateRoomSelectionStates() {
        logger.info("RadiationTab.updateRoomSelectionStates() - Начало обновления состояний комнат");
        int updatedCount = 0;

        if (currentBuilding == null) return;

        for (Floor floor : currentBuilding.getFloors()) {
            for (Space space : floor.getSpaces()) {
                for (Room room : space.getRooms()) {
                    Boolean selected = globalRoomSelectionMap.get(room.getId());
                    if (selected != null) {
                        room.setRadiationSelected(selected);
                        updatedCount++;
                    }
                }
            }
        }
        logger.info("RadiationTab.updateRoomSelectionStates() - Обновлено {} комнат", updatedCount);
        rad_refreshHighlights();
    }

    public void refreshFloors() {
        String selectedFloorNumber = null;
        if (floorList != null && floorList.getSelectedValue() != null) {
            selectedFloorNumber = floorList.getSelectedValue().getNumber();
        }

        floorListModel.clear();
        if (currentBuilding != null) {
            int secIdx = 0;
            if (sectionList != null && sectionList.getSelectedIndex() >= 0) {
                secIdx = sectionList.getSelectedIndex();
            }
            java.util.List<Floor> list = new java.util.ArrayList<>();
            for (Floor f : currentBuilding.getFloors()) {
                if (f.getSectionIndex() == secIdx) list.add(f);
            }
            list.sort(java.util.Comparator.comparingInt(Floor::getPosition));
            list.forEach(floorListModel::addElement);
        }

        boolean restored = false;
        if (selectedFloorNumber != null) {
            for (int i = 0; i < floorListModel.size(); i++) {
                if (floorListModel.get(i).getNumber().equals(selectedFloorNumber)) {
                    floorList.setSelectedIndex(i);
                    restored = true;
                    break;
                }
            }
        }
        if (!restored && !floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        }

        updateSpaceList();
        updateRoomList();
        rad_refreshHighlights();
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


    /** Есть ли в помещении хотя бы одна комната с галочкой радиации? */
    private boolean rad_hasAnyInSpace(Space s) {
        if (s == null) return false;
        for (Room r : s.getRooms()) {
            Boolean v = globalRoomSelectionMap.get(r.getId());
            boolean chosen = (v != null) ? v : r.isRadiationSelected();
            if (chosen) return true;
        }
        return false;
    }

    /** Есть ли на этом этаже хотя бы одно помещение с галочкой радиации? */
    private boolean rad_hasAnyOnFloor(Floor f) {
        if (f == null) return false;
        for (Space s : f.getSpaces()) {
            if (rad_hasAnyInSpace(s)) return true;
        }
        return false;
    }

    /** Перерисовать списки/таблицы с учётом подсветки. */
    private void rad_refreshHighlights() {
        try { if (floorList != null) floorList.repaint(); } catch (Throwable ignore) {}
        try { if (spaceTable != null) spaceTable.repaint(); } catch (Throwable ignore) {}
    }

}