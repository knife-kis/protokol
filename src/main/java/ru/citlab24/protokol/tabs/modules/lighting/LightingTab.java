package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.event.TableModelEvent;
import java.awt.*;
import java.util.List;
import java.util.*;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * Вкладка «Освещение» в стиле «Радиации».
 * Слева направо: Секции → Этажи → Помещения → Комнаты (с чекбоксами).
 * Автопроставление галочек: первый ЖИЛОЙ этаж (RESIDENTIAL) или MIXED с квартирами,
 * исключая санузлы/ванные/душ/туалет/гардероб/коридор. Ручные клики не перетираем.
 */
public final class LightingTab extends JPanel {

    private static final Color PANEL_COLOR = new Color(0,115,200);
    private static final Font HEADER_FONT = UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);

    // Модель здания
    private Building currentBuilding;

    // Списки/таблицы
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private final JList<Section> sectionList = new JList<>(sectionListModel);

    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final JList<Floor> floorList = new JList<>(floorListModel);

    private final SpaceTableModel spaceTableModel = new SpaceTableModel();
    private final JTable spaceTable = new JTable(spaceTableModel);

    private final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final Set<Integer> userTouchedRooms = new HashSet<>();
    private final Set<String> processedSpaceKeys = new HashSet<>();

    private final LightingRoomsTableModel roomsTableModel =
            new LightingRoomsTableModel(globalRoomSelectionMap, new Consumer<Integer>() {
                @Override public void accept(Integer id) { userTouchedRooms.add(id); }
            });
    private final JTable roomTable = new JTable(roomsTableModel);

    private boolean autoApplyRulesOnDisplay = true;

    // ====== Конструктор ======
    public LightingTab(Building building) {
        this.currentBuilding = building;

        setLayout(new BorderLayout(10,10));
        setBorder(BorderFactory.createEmptyBorder(10,10,10,10));
        add(buildMainPanel(), BorderLayout.CENTER);

        // Лиснеры выбора
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setFixedCellHeight(28);
        sectionList.addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) refreshFloors(); });

        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setFixedCellHeight(28);
        floorList.addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) updateSpaceList(); });
        try {
            // Если у тебя есть кастомные рендереры — используем
            floorList.setCellRenderer(new ru.citlab24.protokol.tabs.renderers.FloorListRenderer());
        } catch (Throwable ignore) { /* не критично */ }

        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceTable.getSelectionModel().addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) updateRoomList(); });

        roomTable.getTableHeader().setFont(HEADER_FONT);
        roomTable.setRowHeight(26);
        roomTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        roomsTableModel.addTableModelListener(e -> {
            if (e.getType() == TableModelEvent.UPDATE && e.getColumn() == 0) {
                int row = e.getFirstRow();
                Room r = roomsTableModel.getRoomAt(row);
                if (r != null) userTouchedRooms.add(r.getId());
            }
        });

        // Инициализация
        setBuilding(building, true);
    }

    // ====== Публичные API (как в Радиации) ======
    public void setBuilding(Building building, boolean autoApplyDefaults) {
        this.currentBuilding = building;
        this.autoApplyRulesOnDisplay = autoApplyDefaults;
        processedSpaceKeys.clear();
        userTouchedRooms.clear();
        globalRoomSelectionMap.clear();
        if (autoApplyDefaults) {
            applyAutoOnFirstFloorsAcrossSections(); // ← проставим галочки сразу
        }

        if (building != null) {
            for (Floor f : building.getFloors())
                for (Space s : f.getSpaces())
                    for (Room r : s.getRooms())
                        globalRoomSelectionMap.put(r.getId(), r.isSelected());
        }
        refreshSections();
        refreshFloors(); // внутри вызовется updateSpaceList()
    }

    /** Пробросить состояния чекбоксов обратно в доменную модель. */
    public void updateRoomSelectionStates() {
        if (currentBuilding == null) return;
        for (Floor f : currentBuilding.getFloors())
            for (Space s : f.getSpaces())
                for (Room r : s.getRooms()) {
                    Boolean v = globalRoomSelectionMap.get(r.getId());
                    if (v != null) r.setSelected(v);
                }
    }

    // ====== UI ======
    private JPanel buildMainPanel() {
        JPanel p = new JPanel(new GridBagLayout());
        GridBagConstraints g = new GridBagConstraints();
        g.insets = new Insets(0,0,0,0);
        g.fill = GridBagConstraints.BOTH;
        g.weighty = 1.0;

        g.gridx = 0; g.gridy = 0; g.weightx = 0.25;
        p.add(panelWithBorder("Секции", new JScrollPane(sectionList)), g);

        g.gridx = 1; g.weightx = 0.25;
        p.add(panelWithBorder("Этажи", new JScrollPane(floorList)), g);

        g.gridx = 2; g.weightx = 0.25;
        p.add(panelWithBorder("Помещения", new JScrollPane(spaceTable)), g);

        g.gridx = 3; g.weightx = 0.25;
        p.add(panelWithBorder("Комнаты", new JScrollPane(roomTable)), g);

        return p;
    }

    private JPanel panelWithBorder(String title, JComponent inner) {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(createTitledBorder(title, PANEL_COLOR));
        p.add(inner, BorderLayout.CENTER);
        return p;
    }

    private TitledBorder createTitledBorder(String title, Color color) {
        return BorderFactory.createTitledBorder(null, title, TitledBorder.LEFT, TitledBorder.TOP, HEADER_FONT, color);
    }

    // ====== Refresh ======
    private void refreshSections() {
        sectionListModel.clear();
        if (currentBuilding == null) return;
        List<Section> secs = currentBuilding.getSections(); // порядок = индексы в модели
        for (Section s : secs) sectionListModel.addElement(s);
        if (!sectionListModel.isEmpty() && sectionList.getSelectedIndex() < 0) {
            sectionList.setSelectedIndex(0);
        }
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding == null) { updateSpaceList(); return; }

        final int secIdx = Math.max(0, sectionList.getSelectedIndex()); // <- как в радиации

        List<Floor> list = currentBuilding.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .sorted(Comparator.comparingInt(Floor::getPosition))
                .collect(Collectors.toList());

        for (Floor f : list) floorListModel.addElement(f);
        if (!floorListModel.isEmpty() && floorList.getSelectedIndex() < 0) {
            floorList.setSelectedIndex(0);
        }
        updateSpaceList(); // принудительно, если выбор не изменился
    }

    private int computeRawSectionIndex() {
        if (currentBuilding == null) return 0;
        Section sel = sectionList.getSelectedValue();
        if (sel == null) return 0;
        List<Section> raw = currentBuilding.getSections(); // исходный список (индексы совпадают с floor.sectionIndex)
        int idx = raw.indexOf(sel);
        return (idx >= 0) ? idx : 0;
    }

    private void updateSpaceList() {
        Floor floor = floorList.getSelectedValue();
        if (floor == null) {
            spaceTableModel.setSpaces(Collections.emptyList());
            updateRoomList();
            return;
        }
        List<Space> spaces = new ArrayList<>(floor.getSpaces());
        spaces.sort(Comparator.comparingInt(Space::getPosition));
        spaceTableModel.setSpaces(spaces);

        if (autoApplyRulesOnDisplay) autoApplyForFirstResidentialFloorIfNeeded(floor);

        if (spaceTable.getRowCount() > 0 && spaceTable.getSelectedRow() < 0) {
            spaceTable.setRowSelectionInterval(0,0);
        }
        updateRoomList();
    }

    private void updateRoomList() {
        int row = spaceTable.getSelectedRow();
        if (row < 0) { roomsTableModel.setRooms(Collections.emptyList()); return; }
        Space s = spaceTableModel.getSpaceAt(row);
        if (s == null) { roomsTableModel.setRooms(Collections.emptyList()); return; }
        roomsTableModel.setRooms(new ArrayList<>(s.getRooms()));
    }

    // ====== Автопроставление для первого жилого этажа ======
    private void autoApplyForFirstResidentialFloorIfNeeded(Floor selectedFloor) {
        if (currentBuilding == null) return;

        Map<Integer, Integer> firstBySection = findFirstResidentialFloorPositionBySection(currentBuilding);
        Integer firstPos = firstBySection.get(selectedFloor.getSectionIndex());
        if (firstPos == null) return;
        if (selectedFloor.getPosition() != firstPos) return;

        for (Space s : selectedFloor.getSpaces()) {
            if (s.getType() != Space.SpaceType.APARTMENT) continue;

            String key = spaceKey(selectedFloor, s); // устойчивый ключ (без Space.id)
            if (processedSpaceKeys.contains(key)) continue;

            for (Room r : s.getRooms()) {
                if (isExcludedRoom(r.getName())) continue;
                if (!userTouchedRooms.contains(r.getId())) {
                    globalRoomSelectionMap.put(r.getId(), true);
                }
            }
            processedSpaceKeys.add(key);
        }
        roomTable.repaint();
    }

    /** Жилой этаж: RESIDENTIAL или MIXED при наличии хотя бы одной квартиры. */
    private Map<Integer, Integer> findFirstResidentialFloorPositionBySection(Building bld) {
        Map<Integer, Integer> result = new HashMap<>();
        Map<Integer, List<Floor>> bySec = bld.getFloors().stream().collect(Collectors.groupingBy(Floor::getSectionIndex));
        for (Map.Entry<Integer, List<Floor>> e : bySec.entrySet()) {
            int sec = e.getKey();
            List<Floor> fs = new ArrayList<>(e.getValue());
            fs.sort(Comparator.comparingInt(Floor::getPosition));

            Integer firstPos = null;
            for (Floor f : fs) {
                if (f.getType() == Floor.FloorType.RESIDENTIAL) { firstPos = f.getPosition(); break; }
                if (f.getType() == Floor.FloorType.MIXED) {
                    boolean hasApts = f.getSpaces().stream().anyMatch(s -> s.getType() == Space.SpaceType.APARTMENT);
                    if (hasApts) { firstPos = f.getPosition(); break; }
                }
            }
            if (firstPos != null) result.put(sec, firstPos);
        }
        return result;
    }

    private String spaceKey(Floor f, Space s) {
        int sec = (f != null) ? f.getSectionIndex() : -1;
        int fpos = (f != null) ? f.getPosition() : -1;
        int spos = (s != null) ? s.getPosition() : -1;
        String sid = (s != null && s.getIdentifier() != null) ? s.getIdentifier() : "";
        return sec + "/" + fpos + "/" + spos + "/" + sid;
    }

    private boolean isExcludedRoom(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(Locale.ROOT);
        return n.contains("сануз") || n.contains("сан уз") || n.contains("сан. уз")
                || n.contains("ванн") || n.contains("душ") || n.contains("туал")
                || n.contains("wc") || n.contains("с/у") || n.contains("су")
                || n.contains("гардероб") || n.contains("коридор");
    }

    // ====== Таблица помещений (1 колонка) ======
    static final class SpaceTableModel extends javax.swing.table.AbstractTableModel {
        private final String[] COLUMN_NAMES = {"Название"};
        private final List<Space> spaces = new ArrayList<>();

        void setSpaces(List<Space> newSpaces) {
            spaces.clear();
            if (newSpaces != null) spaces.addAll(newSpaces);
            fireTableDataChanged();
        }
        Space getSpaceAt(int row) {
            return (row >= 0 && row < spaces.size()) ? spaces.get(row) : null;
        }

        @Override public int getRowCount() { return spaces.size(); }
        @Override public int getColumnCount() { return 1; }
        @Override public String getColumnName(int column) { return COLUMN_NAMES[column]; }
        @Override public Object getValueAt(int rowIndex, int columnIndex) {
            Space s = spaces.get(rowIndex);
            return (s != null) ? s.getIdentifier() : "";
        }
    }
    // стало: добавили безопасное обновление UI без автопроставления
    /** Перечитать UI из текущей доменной модели без автопроставления. */
    public void refreshData() {
        boolean prev = this.autoApplyRulesOnDisplay;
        try {
            this.autoApplyRulesOnDisplay = false; // не триггерим авто-галочки при простом рефреше
            refreshSections();
            refreshFloors();   // внутри принудительно перегружает список помещений
        } finally {
            this.autoApplyRulesOnDisplay = prev;
        }
    }
    /** Применить правила ко всему зданию (идемпотентно, не трогаем ручные клики). */
    private void applyAutoOnFirstFloorsAcrossSections() {
        if (currentBuilding == null) return;
        Map<Integer, Integer> firstBySection = findFirstResidentialFloorPositionBySection(currentBuilding);

        for (Floor f : currentBuilding.getFloors()) {
            Integer firstPos = firstBySection.get(f.getSectionIndex());
            if (firstPos == null || f.getPosition() != firstPos) continue;

            for (Space s : f.getSpaces()) {
                if (s.getType() != Space.SpaceType.APARTMENT) continue;
                for (Room r : s.getRooms()) {
                    if (isExcludedRoom(r.getName())) continue;
                    if (!userTouchedRooms.contains(r.getId())) {
                        globalRoomSelectionMap.put(r.getId(), true);
                    }
                }
            }
        }
        if (roomTable != null) roomTable.repaint();
    }

}
