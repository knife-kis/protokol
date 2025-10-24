package ru.citlab24.protokol.tabs.modules.microclimateTab;

import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.SpinnerEditor;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.util.*;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Collectors;

/**
 * Вкладка «Микроклимат» — визуально как «Освещение», но с другой логикой галочек
 * и с дополнительной колонкой «Нар. стены (0..4)» со спиннером.
 */
public class MicroclimateTab extends JPanel {

    // ===== Модель =====
    private Building currentBuilding;

    // Списки/таблицы
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private final JList<Section> sectionList = new JList<>(sectionListModel);

    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final JList<Floor> floorList = new JList<>(floorListModel);

    private final SpaceTableModel spaceTableModel = new SpaceTableModel();
    private final JTable spaceTable = new JTable(spaceTableModel);

    private final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final Set<Integer> userTouchedRooms = new HashSet<>(); // комнаты, по которым юзер кликал вручную
    private final Set<String> processedSpaceKeys = new HashSet<>(); // чтобы не применять авто-правила повторно

    private final MicroclimateRoomsTableModel roomsTableModel =
            new MicroclimateRoomsTableModel(globalRoomSelectionMap, new Consumer<Integer>() {
                @Override public void accept(Integer id) { userTouchedRooms.add(id); }
            });
    private final JTable roomTable = new JTable(roomsTableModel);

    private boolean autoApplyRulesOnDisplay = true;

    // ===== Конструкторы =====
    public MicroclimateTab(Building building) {
        this.currentBuilding = building;
        setLayout(new BorderLayout());
        add(buildMainPanel(), BorderLayout.CENTER);

        // Навешиваем обработчики
        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloors();
        });
        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                updateSpaceList();
                updateRoomList();
                autoApplyForFirstResidentialFloorIfNeeded(floorList.getSelectedValue());
            }
        });
        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) updateRoomList();
        });

        // Инициализация данными
        if (currentBuilding != null) display(currentBuilding, true);
    }

    /** Сохранённый старый API — теперь просто прокидываем building из MainFrame. */
    public MicroclimateTab() {
        this(null);
    }

    // ===== Публичные методы =====
    /** Полная перерисовка вкладки по текущему зданию. autoApplyDefaults = true — применить автологики. */
    public void display(Building bld, boolean autoApplyDefaults) {
        if (bld == null) return;
        this.currentBuilding = bld;

        // 1) Сохраняем старые отмеченные и "ручные" комнаты
        Map<Integer, Boolean> oldMap = new HashMap<>(globalRoomSelectionMap);
        Set<Integer> oldTouched = new HashSet<>(userTouchedRooms);
        Set<Integer> oldIds = new HashSet<>();
        for (Floor f : bld.getFloors())
            for (Space s : f.getSpaces())
                for (Room r : s.getRooms())
                    oldIds.add(roomKey(r));


        // 2) Переносим старые галочки (по id)
        globalRoomSelectionMap.clear();
        for (Floor f : bld.getFloors())
            for (Space s : f.getSpaces())
                for (Room r : s.getRooms()) {
                    int k = roomKey(r);
                    globalRoomSelectionMap.put(k, oldMap.getOrDefault(k, r.isSelected()));
                }

        // 3) Сбрасываем служебные наборы
        userTouchedRooms.clear();
        userTouchedRooms.addAll(oldTouched);
        processedSpaceKeys.clear();

        // 4) Автопроставление ТОЛЬКО для новых комнат
        if (autoApplyDefaults) {
            applyOfficeAutoAlwaysRespectManual();                 // ← всегда для офисов
            applyAutoOnFirstFloorsAcrossSectionsOnlyForNewRooms(oldIds); // жилые первые этажи — как раньше
        }

        // 5) Обновляем UI
        refreshSections();
        refreshFloors();
    }
    public void setBuilding(Building bld, boolean autoApplyDefaults) {
        display(bld, autoApplyDefaults);
    }
    public void refreshData() {
        refreshSections();
        refreshFloors();
    }
    /** Пробросить состояния чекбоксов в доменную модель. */
    public void updateRoomSelectionStates() {
        if (currentBuilding == null) return;
        for (Floor f : currentBuilding.getFloors())
            for (Space s : f.getSpaces())
                for (Room r : s.getRooms()) {
                    Boolean v = globalRoomSelectionMap.get(roomKey(r));
                    if (v != null) r.setSelected(v);
                }
    }

    // ===== Автопроставление: ОФИСЫ =====
    /** Для офисных помещений — отмечаем все комнаты, кроме исключений. Только для НОВЫХ комнат. */
    private void applyOfficeAutoAlwaysRespectManual() {
        if (currentBuilding == null) return;
        for (Floor f : currentBuilding.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (s.getType() != Space.SpaceType.OFFICE) continue;
                for (Room r : s.getRooms()) {
                    int k = roomKey(r);
                    if (userTouchedRooms.contains(k)) continue;
                    boolean excluded = isExcludedRoom(r.getName());
                    globalRoomSelectionMap.put(k, !excluded);
                }
            }
        }
    }

    // ===== Автопроставление: ЖИЛЫЕ (первый этаж каждой секции) =====
    /** Для первой жилой/совмещённой по секции — отметить ВСЕ комнаты в квартирах. Только для НОВЫХ. */
    private void applyAutoOnFirstFloorsAcrossSectionsOnlyForNewRooms(Set<Integer> oldIds) {
        if (currentBuilding == null) return;

        Map<Integer, Integer> firstPosBySection = findFirstResidentialFloorPositionBySection(currentBuilding);

        for (Floor f : currentBuilding.getFloors()) {
            Integer firstPos = firstPosBySection.get(f.getSectionIndex());
            if (firstPos == null || f.getPosition() != firstPos) continue;

            for (Space s : f.getSpaces()) {
                if (s.getType() != Space.SpaceType.APARTMENT) continue;
                for (Room r : s.getRooms()) {
                    int k = roomKey(r);
                    if (oldIds.contains(k)) continue;
                    if (userTouchedRooms.contains(k)) continue;
                    globalRoomSelectionMap.put(k, true);
                }
            }
        }
        if (roomTable != null) roomTable.repaint();
    }

    /** Вызывается при выборе этажа — если это первый жилой/совмещённый, отметим все комнаты в квартирах. */
    private void autoApplyForFirstResidentialFloorIfNeeded(Floor selectedFloor) {
        if (currentBuilding == null || selectedFloor == null) return;

        Map<Integer, Integer> firstBySection = findFirstResidentialFloorPositionBySection(currentBuilding);
        Integer firstPos = firstBySection.get(selectedFloor.getSectionIndex());
        if (firstPos == null) return;
        if (selectedFloor.getPosition() != firstPos) return;

        for (Space s : selectedFloor.getSpaces()) {
            if (s.getType() != Space.SpaceType.APARTMENT) continue;

            String key = spaceKey(selectedFloor, s);
            if (processedSpaceKeys.contains(key)) continue;

            for (Room r : s.getRooms()) {
                int k = roomKey(r);
                if (!userTouchedRooms.contains(k)) {
                    globalRoomSelectionMap.put(k, true);
                }
            }
            processedSpaceKeys.add(key);
        }
        roomTable.repaint();
    }

    /** По клику «Проставить галочки на этаже» — отметить ВСЕ комнаты на выбранном этаже (вне зависимости от типа). */
    private void selectAllRoomsOnSelectedFloor() {
        Floor f = floorList.getSelectedValue();
        if (f == null) return;
        for (Space s : f.getSpaces())
            for (Room r : s.getRooms()) {
                globalRoomSelectionMap.put(roomKey(r), true);
            }
        roomTable.repaint();
    }

    // ===== Вспомогательные правила/поиск =====
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

    private static String spaceKey(Floor f, Space s) {
        int sec = (f != null) ? f.getSectionIndex() : -1;
        int fpos = (f != null) ? f.getPosition() : -1;
        int spos = (s != null) ? s.getPosition() : -1;
        String sid = (s != null && s.getIdentifier() != null) ? s.getIdentifier() : "";
        return sec + "/" + fpos + "/" + spos + "/" + sid;
    }

    private static boolean isExcludedRoom(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(Locale.ROOT);
        return n.contains("сануз") || n.contains("сан уз") || n.contains("сан. уз")
                || n.contains("ванн") || n.contains("душ") || n.contains("туал")
                || n.contains("wc") || n.contains("с/у") || n.contains("су");
    }

    // ===== UI =====
    private JPanel buildMainPanel() {
        JPanel p = new JPanel(new GridBagLayout());
        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(4,4,4,4);
        gc.fill = GridBagConstraints.BOTH;
        gc.weightx = 1; gc.weighty = 1;

        // Секции
        gc.gridx = 0; gc.gridy = 0; gc.weightx = 0.25;
        p.add(createSectionPanel(), gc);

        // Этажи
        gc.gridx = 1; gc.gridy = 0; gc.weightx = 0.25;
        p.add(createFloorPanel(), gc);

        // Помещения и комнаты
        JPanel right = new JPanel(new GridLayout(2, 1, 6, 6));
        right.add(createSpacePanel());
        right.add(createRoomPanel());
        gc.gridx = 2; gc.gridy = 0; gc.weightx = 0.5;
        p.add(right, gc);

        return p;
    }

    private JPanel createSectionPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Секции", new Color(33, 150, 243)));

        sectionList.setCellRenderer(new DefaultListCellRenderer() {
            @Override public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                JLabel lb = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Section) {
                    Section s = (Section) value;
                    lb.setText(s.getName());
                }
                return lb;
            }
        });

        panel.add(new JScrollPane(sectionList), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createFloorPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Этажи", new Color(76, 175, 80)));
        floorList.setCellRenderer(new DefaultListCellRenderer() {
            @Override public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                JLabel lb = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Floor) {
                    Floor f = (Floor) value;
                    lb.setText(f.getNumber() + " — " + String.valueOf(f.getType()));
                }
                return lb;
            }
        });

        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT));

        JButton selectAllOnFloorBtn = new JButton("Проставить галочки на этаже");
        selectAllOnFloorBtn.addActionListener(e -> selectAllRoomsOnSelectedFloor());
        btns.add(selectAllOnFloorBtn);

        // ⬇⬇⬇ ДОБАВЛЕННАЯ КНОПКА ЭКСПОРТА ⬇⬇⬇
        JButton exportBtn = new JButton("Экспорт (микроклимат)");
        exportBtn.addActionListener(e -> {
            // экспорт по выбранной секции; если нужно «все секции» — вместо computeRawSectionIndex() поставь -1
            int sectionIndex = computeRawSectionIndex();
            MicroclimateExcelExporter.export(currentBuilding, sectionIndex, MicroclimateTab.this);
        });
        btns.add(exportBtn);

        panel.add(new JScrollPane(floorList), BorderLayout.CENTER);
        panel.add(btns, BorderLayout.SOUTH);
        return panel;
    }


    private JPanel createSpacePanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Помещения", new Color(255, 152, 0)));
        spaceTable.setFillsViewportHeight(true);
        spaceTable.setRowHeight(26);
        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }

    private JPanel createRoomPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(createTitledBorder("Комнаты", new Color(156, 39, 176)));

        roomTable.setFillsViewportHeight(true);
        roomTable.setRowHeight(26);

        // Спиннер для «Нар. стены» (0..4, шаг 1)
        roomTable.getColumnModel().getColumn(2).setCellRenderer(new WallsButtonsRenderer());
        roomTable.getColumnModel().getColumn(2).setCellEditor(new WallsButtonsEditor());

// опционально сделаем столбец шире
        roomTable.getColumnModel().getColumn(2).setPreferredWidth(160);

        // Ширины столбцов
        roomTable.getColumnModel().getColumn(0).setPreferredWidth(90);
        roomTable.getColumnModel().getColumn(1).setPreferredWidth(260);
        roomTable.getColumnModel().getColumn(2).setPreferredWidth(110);

        panel.add(new JScrollPane(roomTable), BorderLayout.CENTER);
        return panel;
    }

    private static TitledBorder createTitledBorder(String title, Color color) {
        TitledBorder tb = BorderFactory.createTitledBorder(title);
        tb.setTitleFont(new Font("Segoe UI", Font.BOLD, 14));
        tb.setTitleColor(color);
        return tb;
    }

    // ===== Рефрешы =====
    private void refreshSections() {
        sectionListModel.clear();
        if (currentBuilding == null) return;

        // Порядок секций строго по position
        List<Section> secs = new ArrayList<>(currentBuilding.getSections());
        secs.sort(SECTION_ORDER);

        for (Section s : secs) sectionListModel.addElement(s);
        if (!sectionListModel.isEmpty() && sectionList.getSelectedIndex() < 0) {
            sectionList.setSelectedIndex(0);
        }
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding == null) return;

        int rawSec = computeRawSectionIndex();
        List<Floor> floors = new ArrayList<>();
        for (Floor f : currentBuilding.getFloors()) {
            if (rawSec < 0 || f.getSectionIndex() == rawSec) floors.add(f);
        }
        floors.sort(Comparator.comparingInt(Floor::getPosition));
        for (Floor f : floors) floorListModel.addElement(f);

        if (!floors.isEmpty()) floorList.setSelectedIndex(0);
        updateSpaceList();
        updateRoomList();
    }

    private void updateSpaceList() {
        int row = floorList.getSelectedIndex();
        if (row < 0) { spaceTableModel.setSpaces(Collections.emptyList()); return; }
        Floor f = floorListModel.get(row);
        List<Space> spaces = new ArrayList<>(f.getSpaces());
        spaces.sort(SPACE_ORDER);
        spaceTableModel.setSpaces(spaces);
        if (!spaces.isEmpty()) spaceTable.setRowSelectionInterval(0,0);
    }

    private void updateRoomList() {
        int row = spaceTable.getSelectedRow();
        if (row < 0) { roomsTableModel.setRooms(Collections.emptyList()); return; }
        Space s = spaceTableModel.getSpaceAt(row);
        if (s == null) { roomsTableModel.setRooms(Collections.emptyList()); return; }
        List<Room> rooms = new ArrayList<>(s.getRooms());
        rooms.sort(ROOM_ORDER);
        roomsTableModel.setRooms(rooms);
    }

    private int computeRawSectionIndex() {
        if (currentBuilding == null) return 0;
        Section sel = sectionList.getSelectedValue();
        if (sel == null) return 0;
        List<Section> raw = currentBuilding.getSections(); // индексы совпадают с floor.sectionIndex
        int idx = raw.indexOf(sel);
        return (idx >= 0) ? idx : 0;
    }

    // ===== Модели/сортировки =====
    static final Comparator<Section> SECTION_ORDER = Comparator
            .comparingInt(Section::getPosition)
            .thenComparing(Section::getName, Comparator.nullsLast(String::compareToIgnoreCase));

    static final Comparator<Room> ROOM_ORDER = Comparator
            .comparingInt(Room::getPosition)
            .thenComparing(Room::getName, Comparator.nullsLast(String::compareToIgnoreCase));

    static final Comparator<Space> SPACE_ORDER = Comparator
            .comparingInt(Space::getPosition)
            .thenComparing(Space::getIdentifier, Comparator.nullsLast(String::compareToIgnoreCase));

    /** Таблица помещений (1 колонка «Название») — как в «Освещении». */
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
        @Override public Class<?> getColumnClass(int columnIndex) { return String.class; }
        @Override public Object getValueAt(int rowIndex, int columnIndex) {
            Space s = spaces.get(rowIndex);
            return s.getIdentifier();
        }
    }
    // Сегментные кнопки 0..4 — общий помощник
    private static JPanel buildWallsPanel(int value, boolean interactive, Consumer<Integer> onSelect) {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 2, 2));
        ButtonGroup g = new ButtonGroup();

        for (int i = 0; i <= 4; i++) {
            JToggleButton b = new JToggleButton(String.valueOf(i));
            b.setFocusable(false);
            b.setMargin(new Insets(2, 8, 2, 8));

            // FlatLaf: сегментные кнопки + правильная позиция сегмента
            b.putClientProperty("JButton.buttonType", "segmented");
            b.putClientProperty("JButton.segmentPosition",
                    (i == 0) ? "first" : (i == 4) ? "last" : "middle");

            // выразительный selected-стиль
            Color accent = UIManager.getColor("Component.accentColor");
            if (accent == null) accent = new Color(0x2962FF);
            Color onAccent = Color.WHITE;
            Color outline = new Color(0xB0B0B0);

            final Color accent_f = accent;
            final Color onAccent_f = onAccent;
            final Color outline_f = outline;

            b.getModel().addChangeListener(e -> {
                boolean sel = b.isSelected();
                boolean rollover = b.getModel().isRollover();
                b.setForeground(sel ? onAccent_f : UIManager.getColor("Label.foreground"));
                b.setBackground(sel ? accent_f : UIManager.getColor("Button.background"));
                b.setBorder(BorderFactory.createLineBorder(sel ? accent_f.darker()
                        : (rollover ? outline_f.darker() : outline_f)));
                b.setOpaque(true);
                b.setContentAreaFilled(true);
            });

            b.setSelected(i == value);

            if (interactive) {
                int v = i;
                b.addActionListener(e -> { if (onSelect != null) onSelect.accept(v); });
                b.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
                b.setToolTipText("Наружных стен: " + i);
            } else {
                // В рендерере оставляем enabled=true, чтобы не было «серого»
                b.setRequestFocusEnabled(false);
            }

            g.add(b);
            p.add(b);
        }
        return p;
    }

    // ===== Renderer =====
    private final class WallsButtonsRenderer implements javax.swing.table.TableCellRenderer {
        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean isSelected, boolean hasFocus,
                                                       int row, int column) {
            int v = (value instanceof Number) ? ((Number) value).intValue() : 0;
            JPanel p = buildWallsPanel(v, /*interactive=*/false, null);
            p.setBackground(isSelected ? table.getSelectionBackground() : table.getBackground());
            return p;
        }
    }

    // ===== Editor =====
    private final class WallsButtonsEditor extends AbstractCellEditor implements javax.swing.table.TableCellEditor {
        private int current = 0;
        private JPanel panel;

        @Override public Object getCellEditorValue() { return current; }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value,
                                                     boolean isSelected, int row, int column) {
            current = (value instanceof Number) ? ((Number) value).intValue() : 0;
            panel = buildWallsPanel(current, true, v -> { current = v; stopCellEditing(); });

            // клавиатура: ←/→
            panel.getInputMap(JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT)
                    .put(KeyStroke.getKeyStroke("LEFT"), "dec");
            panel.getActionMap().put("dec", new AbstractAction() {
                @Override public void actionPerformed(java.awt.event.ActionEvent e) {
                    if (current > 0) { current--; stopCellEditing(); }
                }
            });
            panel.getInputMap(JComponent.WHEN_ANCESTOR_OF_FOCUSED_COMPONENT)
                    .put(KeyStroke.getKeyStroke("RIGHT"), "inc");
            panel.getActionMap().put("inc", new AbstractAction() {
                @Override public void actionPerformed(java.awt.event.ActionEvent e) {
                    if (current < 4) { current++; stopCellEditing(); }
                }
            });

            return panel;
        }
    }
    // Стабильный ключ комнаты для карт галочек: работает до и после сохранения
    private static int roomKey(Room r) {
        if (r == null) return 0;
        if (r.getId() > 0) return r.getId();
        Integer orig = r.getOriginalRoomId();
        if (orig != null && orig != 0) return orig;
        return System.identityHashCode(r);
    }

}
