package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * Вкладка «Освещение» в стиле «Радиации»:
 * слева — Секции и Этажи, посередине — список Помещений, справа — таблица Комнат (чекбоксы).
 * Автопроставление галочек на первом жилом этаже каждой секции (квартиры), с исключениями.
 */
public class LightingTab extends JPanel {

    private static final Color FLOOR_PANEL_COLOR = new Color(0, 115, 200);
    private static final Font HEADER_FONT = UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);

    private JList<Section> sectionList;
    private DefaultListModel<Section> sectionListModel = new DefaultListModel<>();

    private JList<Floor> floorList;
    private DefaultListModel<Floor> floorListModel = new DefaultListModel<>();

    private JTable spaceTable;
    private SpaceTableModel spaceTableModel = new SpaceTableModel();

    private JTable roomTable;
    private LightingRoomsTableModel roomsTableModel;

    private Building currentBuilding;
    private boolean autoApplyRulesOnDisplay = true;

    // Глобальные состояния чекбоксов по id комнаты (как в радиации)
    public final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final Set<Integer> processedSpaces = new HashSet<>(); // чтобы не перетирать ручной выбор

    public LightingTab(Building building) {
        roomsTableModel = new LightingRoomsTableModel(globalRoomSelectionMap);

        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel mainPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridy = 0;
        gbc.insets = new Insets(0,0,0,0);
        gbc.fill = GridBagConstraints.BOTH;
        gbc.weighty = 1.0;

        // Секции
        gbc.gridx = 0; gbc.weightx = 0.25;
        mainPanel.add(createSectionPanel(), gbc);

        // Этажи
        gbc.gridx = 1; gbc.weightx = 0.25;
        mainPanel.add(createFloorPanel(), gbc);

        // Помещения
        gbc.gridx = 2; gbc.weightx = 0.25;
        mainPanel.add(createSpacePanel(), gbc);

        // Комнаты
        gbc.gridx = 3; gbc.weightx = 0.25;
        mainPanel.add(createRoomPanel(), gbc);

        add(mainPanel, BorderLayout.CENTER);
        setBuilding(building, /*autoApplyDefaults=*/true);
    }

    // ==== Публичные API, как у радиации ====

    public void setBuilding(Building building, boolean autoApplyDefaults) {
        this.currentBuilding = building;
        this.autoApplyRulesOnDisplay = autoApplyDefaults;
        processedSpaces.clear();

        // переносим состояния из модели
        globalRoomSelectionMap.clear();
        if (building != null) {
            for (Floor f : building.getFloors()) {
                for (Space s : f.getSpaces()) {
                    for (Room r : s.getRooms()) {
                        globalRoomSelectionMap.put(r.getId(), r.isSelected());
                    }
                }
            }
        }

        refreshSections();
        refreshFloors();
    }

    public void refreshData() {
        refreshSections();
        refreshFloors();
    }

    /** Сохранить чекбоксы обратно в Room.isSelected() (зови перед сохранением проекта) */
    public void updateRoomSelectionStates() {
        if (currentBuilding == null) return;
        for (Floor f : currentBuilding.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    Boolean sel = globalRoomSelectionMap.get(r.getId());
                    if (sel != null) r.setSelected(sel);
                }
            }
        }
    }

    /** Сохраняем/восстанавливаем снимок состояний (по аналогии с Радиацией) */
    public Map<String, Boolean> saveSelections() {
        Map<String, Boolean> map = new HashMap<>();
        for (Map.Entry<Integer, Boolean> e : globalRoomSelectionMap.entrySet()) {
            map.put(String.valueOf(e.getKey()), e.getValue());
        }
        return map;
    }
    public void restoreSelections(Map<String, Boolean> saved) {
        for (Map.Entry<String, Boolean> e : saved.entrySet()) {
            try {
                globalRoomSelectionMap.put(Integer.parseInt(e.getKey()), e.getValue());
            } catch (NumberFormatException ignore) {}
        }
        updateRoomList();
    }

    // ==== UI блоки ====

    private JPanel createSectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Секции", FLOOR_PANEL_COLOR));

        sectionList = new JList<>(sectionListModel);
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setFixedCellHeight(28);
        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloors();
        });

        panel.add(new JScrollPane(sectionList), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createFloorPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Этажи", FLOOR_PANEL_COLOR));

        floorList = new JList<>(floorListModel);
        floorList.setCellRenderer(new ru.citlab24.protokol.tabs.renderers.FloorListRenderer());
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setFixedCellHeight(28);

        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) updateSpaceList();
        });

        panel.add(new JScrollPane(floorList), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createSpacePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения", FLOOR_PANEL_COLOR));

        spaceTable = new JTable(spaceTableModel);
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) updateRoomList();
        });

        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createRoomPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Комнаты", FLOOR_PANEL_COLOR));

        roomTable = new JTable(roomsTableModel);
        roomTable.getTableHeader().setFont(HEADER_FONT);
        roomTable.setRowHeight(26);
        roomTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        panel.add(new JScrollPane(roomTable), BorderLayout.CENTER);
        return panel;
    }

    // ==== Refresh ====

    private void refreshSections() {
        sectionListModel.clear();
        if (currentBuilding == null) return;
        List<Section> secs = new ArrayList<>(currentBuilding.getSections());
        secs.sort(Comparator.comparingInt(Section::getPosition));
        secs.forEach(sectionListModel::addElement);
        if (!sectionListModel.isEmpty() && sectionList.getSelectedIndex() < 0) {
            sectionList.setSelectedIndex(0);
        }
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding == null) return;

        int secIdx = Math.max(0, sectionList.getSelectedIndex());
        List<Floor> list = currentBuilding.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .sorted(Comparator.comparingInt(Floor::getPosition))
                .collect(Collectors.toList());

        list.forEach(floorListModel::addElement);
        if (!floorListModel.isEmpty() && floorList.getSelectedIndex() < 0) {
            floorList.setSelectedIndex(0);
        }
    }

    private void updateSpaceList() {
        spaceTableModel.clear();
        Floor selectedFloor = floorList.getSelectedValue();
        if (selectedFloor == null) return;

        for (Space s : selectedFloor.getSpaces()) {
            spaceTableModel.addSpace(s);
        }

        if (autoApplyRulesOnDisplay) {
            autoApplyForFirstResidentialFloorIfNeeded(selectedFloor);
        }

        // Выбрать первую строку (если не выбрано)
        if (spaceTable.getRowCount() > 0 && spaceTable.getSelectedRow() < 0) {
            spaceTable.setRowSelectionInterval(0, 0);
        }

        updateRoomList();
    }

    private void updateRoomList() {
        int row = spaceTable.getSelectedRow();
        if (row < 0) {
            roomsTableModel.setRooms(List.of());
            return;
        }
        Space s = spaceTableModel.getSpaceAt(row);
        if (s == null) {
            roomsTableModel.setRooms(List.of());
            return;
        }
        roomsTableModel.setRooms(new ArrayList<>(s.getRooms()));
    }

    // ==== Логика автопроставления ====

    private void autoApplyForFirstResidentialFloorIfNeeded(Floor selectedFloor) {
        // Берём карту "секция -> позиция первого жилого этажа"
        Map<Integer, Integer> firstBySection = findFirstResidentialFloorPositionBySection(currentBuilding);

        Integer firstPos = firstBySection.get(selectedFloor.getSectionIndex());
        if (firstPos == null) return;
        if (selectedFloor.getPosition() != firstPos) return;

        // На первом жилом этаже: отмечаем ВСЕ квартиры (кроме исключений по комнатам)
        for (Space s : selectedFloor.getSpaces()) {
            if (s.getType() == Space.SpaceType.APARTMENT && !processedSpaces.contains(s.getId())) {
                for (Room r : s.getRooms()) {
                    boolean excluded = isExcludedRoom(r.getName());
                    // если пользователь уже явно трогал — не перезатираем
                    globalRoomSelectionMap.putIfAbsent(r.getId(), !excluded);
                    if (!globalRoomSelectionMap.containsKey(r.getId())) {
                        globalRoomSelectionMap.put(r.getId(), !excluded);
                    }
                }
                processedSpaces.add(s.getId());
            }
        }
    }

    /**
     * Жилой этаж: FloorType.RESIDENTIAL ИЛИ FloorType.MIXED, но только если на нём есть квартиры.
     * Берём самый верхний в списке (по position) — у тебя position уже задаёт порядок.
     */
    private Map<Integer, Integer> findFirstResidentialFloorPositionBySection(Building bld) {
        Map<Integer, Integer> result = new HashMap<>();
        if (bld == null) return result;

        Map<Integer, List<Floor>> bySec = bld.getFloors().stream()
                .collect(Collectors.groupingBy(Floor::getSectionIndex));

        for (Map.Entry<Integer, List<Floor>> e : bySec.entrySet()) {
            int sec = e.getKey();
            List<Floor> fs = new ArrayList<>(e.getValue());
            fs.sort(Comparator.comparingInt(Floor::getPosition)); // «сверху тот и жилой»

            Integer firstPos = null;
            for (Floor f : fs) {
                if (f.getType() == Floor.FloorType.RESIDENTIAL) {
                    firstPos = f.getPosition();
                    break;
                }
                if (f.getType() == Floor.FloorType.MIXED) {
                    boolean hasApts = f.getSpaces().stream()
                            .anyMatch(s -> s.getType() == Space.SpaceType.APARTMENT);
                    if (hasApts) {
                        firstPos = f.getPosition();
                        break;
                    }
                }
            }
            if (firstPos != null) result.put(sec, firstPos);
        }
        return result;
    }

    private boolean isExcludedRoom(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(Locale.ROOT);
        return n.contains("сануз") || n.contains("сан уз") || n.contains("сан. уз")
                || n.contains("ванн") || n.contains("душ") || n.contains("туал")
                || n.contains("wc") || n.contains("с/у") || n.contains("су")
                || n.contains("гардероб") || n.contains("коридор");
    }

    private TitledBorder createTitledBorder(String title, Color color) {
        return BorderFactory.createTitledBorder(
                null, title, TitledBorder.LEFT, TitledBorder.TOP, HEADER_FONT, color);
    }

    // ==== Space table model (как в радиации: одна колонка «Название») ====

    static class SpaceTableModel extends AbstractTableModel {
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
        public Space getSpaceAt(int row) {
            return (row >= 0 && row < spaces.size()) ? spaces.get(row) : null;
        }

        @Override public int getRowCount() { return spaces.size(); }
        @Override public int getColumnCount() { return COLUMN_NAMES.length; }
        @Override public String getColumnName(int column) { return COLUMN_NAMES[column]; }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            Space space = spaces.get(rowIndex);
            return space.getIdentifier();
        }
    }
}
