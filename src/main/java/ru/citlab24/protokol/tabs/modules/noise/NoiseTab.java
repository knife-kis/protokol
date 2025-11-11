package ru.citlab24.protokol.tabs.modules.noise;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableCellRenderer;
import java.awt.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Method;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Вкладка «Шумы».
 * Слева направо: Секции (30%) → Этажи (30%) → Помещения (20%) → Таблица комнат (50%).
 * В таблице: [чекбокс «Измер.» | имя комнаты | столбец источников (набор независимых тумблеров)].
 *
 * НИЧЕГО не добавляем в модели Room/Space/Floor — все состояния храним по ключу:
 *   sectionIndex|floorNumber|spaceIdentifier|roomName
 */
public class NoiseTab extends JPanel {


    private Building building;

    // Модели списков
    private final DefaultListModel<Section> sectionModel = new DefaultListModel<>();
    private final DefaultListModel<Floor>   floorModel   = new DefaultListModel<>();
    private final DefaultListModel<Space>   spaceModel   = new DefaultListModel<>();
    private NoiseFilter filter;
    // Компоненты
    private JList<Section> sectionList;
    private JList<Floor>   floorList;
    private JList<Space>   spaceList;
    private JToggleButton[] filterBtns;
    private JTable         roomsTable;

    // Табличная модель комнат
    private NoiseRoomsTableModel tableModel;

    // Снимок состояний по ключу (совместим с DatabaseManager)
    // Ключ: sectionIndex|этаж|помещение|комната
    private final Map<String, DatabaseManager.NoiseValue> byKey = new LinkedHashMap<>();

    // Короткие подписи для источников (влезают в ячейку)
    private static final String[] SRC_SHORT = {"Лифт","Вент","Завеса","ИТП","ПНС","Э/Щ","Авто","Зум"};

    public NoiseTab(Building building) {
        this.building = (building != null) ? building : new Building();
        buildUI();
        refreshData();
    }

    public NoiseTab() {
        this(null);
    }

    /* ================== ПУБЛИЧНЫЕ API, которые вызывает BuildingTab ================== */

    public void setBuilding(Building b) {
        this.building = (b != null) ? b : new Building();
    }

    /** Обновить все списки по текущему building. */
    public void refreshData() {
        applyGlobalFilter(); // учитывает текущее состояние кнопок фильтра
    }


    /** Применить сохранённые настройки к текущему зданию (в память вкладки). */
    public void applySelectionsByKey(java.util.Map<String, DatabaseManager.NoiseValue> from) {
        if (from == null) return;
        byKey.clear(); // byKey final — только очищаем и заполняем заново

        for (Map.Entry<String, DatabaseManager.NoiseValue> e : from.entrySet()) {
            DatabaseManager.NoiseValue v = (e.getValue() != null) ? e.getValue() : new DatabaseManager.NoiseValue();
            // правило: «Измерение» активно, если включён хотя бы один источник
            v.measure = v.lift || v.vent || v.heatCurtain || v.itp || v.pns || v.electrical || v.autoSrc || v.zum;
            byKey.put(e.getKey(), v);
        }
    }

    /** Сохранить внутренние состояния в БД: собрать актуальные значения из таблицы. */
    /** Снимок состояний по ключу "sectionIndex|floorNumber|spaceId|roomName". */
    public java.util.Map<String, DatabaseManager.NoiseValue> saveSelectionsByKey() {
        // на всякий — зафиксируем активный редактор
        updateRoomSelectionStates();

        java.util.Map<String, DatabaseManager.NoiseValue> snap = new java.util.LinkedHashMap<>();
        if (building == null) return snap;

        for (Floor f : building.getFloors()) {
            String floorNum = (f.getNumber() == null) ? "" : f.getNumber().trim();
            int sec = Math.max(0, f.getSectionIndex());
            for (Space s : f.getSpaces()) {
                String spaceId = (s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
                for (Room r : s.getRooms()) {
                    String roomName = (r.getName() == null) ? "" : r.getName().trim();
                    String key = sec + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    ru.citlab24.protokol.db.DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null) {
                        nv = new ru.citlab24.protokol.db.DatabaseManager.NoiseValue();
                    }

                    // правило: measure = выбран хотя бы один источник
                    boolean any = nv.lift || nv.vent || nv.heatCurtain || nv.itp || nv.pns || nv.electrical || nv.autoSrc || nv.zum;
                    nv.measure = any;

                    snap.put(key, nv);
                }
            }
        }
        return snap;
    }

    /** Зафиксировать текущее редактирование и протолкнуть значения из редактора в модель. */
    public void updateRoomSelectionStates() {
        try {
            if (roomsTable != null && roomsTable.isEditing()) {
                try { roomsTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        } catch (Throwable ignore) {
            // ничего не ломаем, просто не мешаем сохранению проекта
        }
    }


    /* ================== ВНУТРЕННЕЕ: UI ================== */

    private void buildUI() {
        setLayout(new BorderLayout());

        // ===== Верхняя панель фильтров по источникам =====
        JPanel filters = buildFilterPanel();
        add(filters, BorderLayout.NORTH);

        // ==== Списки слева: секции и этажи ====
        sectionList = new JList<>(sectionModel);
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        floorList = new JList<>(floorModel);
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setCellRenderer(new FloorListRenderer());

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                refreshFloors();
                refreshSpaces();
                refreshRooms();
            }
        });
        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                refreshSpaces();
                refreshRooms();
            }
        });

        // ==== Список помещений ====
        spaceList = new JList<>(spaceModel);
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.setCellRenderer(new SpaceListRenderer());
        spaceList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshRooms();
        });

        // ==== Таблица комнат (2 столбца: Комната | Источник) ====
        tableModel = new NoiseRoomsTableModel();
        roomsTable = new JTable(tableModel);
        roomsTable.setRowHeight(28);
        roomsTable.getColumnModel().getColumn(0).setPreferredWidth(220); // Комната
        roomsTable.getColumnModel().getColumn(1).setPreferredWidth(650); // Источник(тумблеры)

        // Рендер/редактор для колонки источников
        NoiseSourcesCell cell = new NoiseSourcesCell(SRC_SHORT);
        roomsTable.getColumnModel().getColumn(1).setCellRenderer(cell);
        roomsTable.getColumnModel().getColumn(1).setCellEditor(cell);

        // ==== Макет с долями ширины ====
        JSplitPane left = new JSplitPane(
                JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(sectionList),
                new JScrollPane(floorList)
        );
        left.setResizeWeight(0.5);

        JSplitPane right = new JSplitPane(
                JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(spaceList),
                new JScrollPane(roomsTable)
        );

        JSplitPane main = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, left, right);
        main.setResizeWeight(0.6);

        add(main, BorderLayout.CENTER);
    }
    private JPanel buildFilterPanel() {
        JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 6));
        p.add(new JLabel("Фильтр по источникам:"));
        filterBtns = new JToggleButton[SRC_SHORT.length];
        for (int i = 0; i < SRC_SHORT.length; i++) {
            JToggleButton b = new JToggleButton(SRC_SHORT[i]);
            b.setFocusable(false);
            b.setMargin(new Insets(2,6,2,6));
            b.addItemListener(e -> applyGlobalFilter()); // глобальная фильтрация
            filterBtns[i] = b;
            p.add(b);
        }
        JButton clear = new JButton("Сброс");
        clear.setFocusable(false);
        clear.addActionListener(e -> {
            for (JToggleButton b : filterBtns) b.setSelected(false);
            applyGlobalFilter();
        });
        p.add(clear);
        return p;
    }

    private Set<String> getActiveFilterSources() {
        Set<String> s = new LinkedHashSet<>();
        if (filterBtns != null) {
            for (JToggleButton b : filterBtns) if (b.isSelected()) s.add(b.getText());
        }
        return s;
    }

    /* ================== ВНУТРЕННЕЕ: наполнение списков ================== */

    private void refreshRooms() {
        commitEditors();
        Space s = spaceList.getSelectedValue();
        if (s == null) { tableModel.setRooms(Collections.emptyList()); return; }

        int secIdx = getSelectedSectionIndex();
        Floor f = floorList.getSelectedValue();
        List<Room> rooms = (filter == null)
                ? s.getRooms().stream()
                .sorted(Comparator.comparingInt(Room::getPosition))
                .collect(Collectors.toList())
                : filter.filterRooms(secIdx, f, s);

        tableModel.setRooms(rooms);
    }

    private void refreshFloors() {
        floorModel.clear();
        int secIdx = getSelectedSectionIndex();
        List<Floor> floors = (filter == null)
                ? building.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .sorted(Comparator.comparingInt(Floor::getPosition))
                .collect(Collectors.toList())
                : filter.filterFloors(secIdx);

        floors.forEach(floorModel::addElement);
        if (!floorModel.isEmpty()) floorList.setSelectedIndex(0);
        else spaceModel.clear();
    }
    private void refreshSpaces() {
        spaceModel.clear();
        Floor f = floorList.getSelectedValue();
        if (f == null) return;

        int secIdx = getSelectedSectionIndex();
        List<Space> spaces = (filter == null)
                ? f.getSpaces().stream()
                .sorted(Comparator.comparingInt(Space::getPosition))
                .collect(Collectors.toList())
                : filter.filterSpaces(secIdx, f);

        spaces.forEach(spaceModel::addElement);
        if (!spaceModel.isEmpty()) spaceList.setSelectedIndex(0);
        else tableModel.setRooms(Collections.emptyList());
    }

    private void refreshSections() {
        sectionModel.clear();
        List<Section> show = (filter == null) ? building.getSections() : filter.filterSections();
        for (Section s : show) sectionModel.addElement(s);
        if (!sectionModel.isEmpty()) sectionList.setSelectedIndex(0);
    }


    private void commitEditors() {
        if (roomsTable == null) return;
        if (roomsTable.isEditing()) {
            try { roomsTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
        }
    }

    /* ================== Табличная модель ================== */

    private final class NoiseRoomsTableModel extends AbstractTableModel {
        private final String[] COLS = {"Комната", "Источник"};
        private List<Room> rows = new ArrayList<>();

        void setRooms(List<Room> rooms) {
            this.rows = (rooms != null) ? rooms : new ArrayList<>();
            fireTableDataChanged();
        }

        @Override public int getRowCount() { return rows.size(); }
        @Override public int getColumnCount() { return COLS.length; }
        @Override public String getColumnName(int column) { return COLS[column]; }
        @Override public Class<?> getColumnClass(int col) {
            return (col == 0) ? String.class : Object.class;
        }
        @Override public boolean isCellEditable(int row, int col) { return col == 1; }

        @Override public Object getValueAt(int row, int col) {
            Room r = rows.get(row);
            String key = keyFor(r);
            DatabaseManager.NoiseValue nv = byKey.get(key);
            Set<String> sources = (nv != null) ? getNvSources(nv) : Collections.emptySet();

            if (col == 0) return r.getName();
            // col == 1
            return sources; // набор активных тумблеров
        }

        @Override public void setValueAt(Object aValue, int row, int col) {
            if (col != 1) return;

            Room r = rows.get(row);
            String key = keyFor(r);
            DatabaseManager.NoiseValue nv = byKey.get(key);
            if (nv == null) {
                nv = newNoiseValue(false, new LinkedHashSet<>());
                byKey.put(key, nv);
            }

            if (aValue instanceof Set) {
                @SuppressWarnings("unchecked")
                Set<String> s = new LinkedHashSet<>((Set<String>) aValue);
                setNvSources(nv, s);
                // «Измерение» = выбран хотя бы один источник
                setNvSelected(nv, !s.isEmpty());

                // ВАЖНО: не вызываем refreshRooms() немедленно (это ведёт к ре-энтерации через stopCellEditing).
                // Делаем отложенное обновление после завершения редактирования.
                if (!getActiveFilterSources().isEmpty()) {
                    SwingUtilities.invokeLater(() -> {
                        if (roomsTable != null && !roomsTable.isEditing()) {
                            // Перерисуем только таблицу комнат — без коммита редактора
                            refreshRooms();
                        }
                    });
                }
            }
        }

        private String keyFor(Room r) {
            Floor f = floorList.getSelectedValue();
            Space s = spaceList.getSelectedValue();
            int secIdx = getSelectedSectionIndex(); // корректный индекс секции в модели Building

            String floorNum = (f == null || f.getNumber() == null) ? "" : f.getNumber().trim();
            String spaceId  = (s == null || s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
            String roomName = (r == null || r.getName() == null) ? "" : r.getName().trim();

            return secIdx + "|" + floorNum + "|" + spaceId + "|" + roomName;
        }

    }


    /* ================== Ячейка с набором тумблеров источников ================== */

    private static final class NoiseSourcesCell extends AbstractCellEditor
            implements TableCellRenderer, TableCellEditor {

        private final String[] labels;
        private JPanel panel;

        NoiseSourcesCell(String[] labels) {
            this.labels = labels.clone();
        }

        private JPanel buildPanel(Set<String> active) {
            JPanel p = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));
            for (String lbl : labels) {
                JToggleButton b = new JToggleButton(lbl);
                b.setMargin(new Insets(2,6,2,6));
                b.setFocusable(false);
                b.setSelected(active != null && active.contains(lbl));
                b.addActionListener(e -> b.setSelected(b.isSelected()));
                p.add(b);
            }
            return p;
        }

        private Set<String> collect(JPanel p) {
            Set<String> s = new LinkedHashSet<>();
            for (Component c : p.getComponents()) {
                if (c instanceof JToggleButton b && b.isSelected()) {
                    s.add(b.getText());
                }
            }
            return s;
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean isSelected, boolean hasFocus,
                                                       int row, int column) {
            @SuppressWarnings("unchecked")
            Set<String> active = (value instanceof Set) ? (Set<String>) value : Collections.emptySet();
            JPanel p = buildPanel(active);
            p.setBackground(isSelected ? table.getSelectionBackground() : table.getBackground());
            return p;
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value,
                                                     boolean isSelected, int row, int column) {
            @SuppressWarnings("unchecked")
            Set<String> active = (value instanceof Set) ? (Set<String>) value : Collections.emptySet();
            panel = buildPanel(active);
            return panel;
        }

        @Override
        public Object getCellEditorValue() {
            return (panel == null) ? Collections.emptySet() : collect(panel);
        }
    }

    /* ================== Рефлексивная обёртка над DatabaseManager.NoiseValue ================== */

    private static DatabaseManager.NoiseValue newNoiseValue(boolean selected, Set<String> sources) {
        DatabaseManager.NoiseValue nv = new DatabaseManager.NoiseValue();
        nv.measure = selected;
        setNvSources(nv, sources);
        return nv;
    }

    @SuppressWarnings("unchecked")
    private static Set<String> getNvSources(DatabaseManager.NoiseValue nv) {
        Set<String> s = new LinkedHashSet<>();
        if (nv.lift)        s.add("Лифт");
        if (nv.vent)        s.add("Вент");
        if (nv.heatCurtain) s.add("Завеса");
        if (nv.itp)         s.add("ИТП");
        if (nv.pns)         s.add("ПНС");
        if (nv.electrical)  s.add("Э/Щ");
        if (nv.autoSrc)     s.add("Авто");
        if (nv.zum)         s.add("Зум");
        return s;
    }


    private static void setNvSelected(DatabaseManager.NoiseValue nv, boolean val) {
        nv.measure = val;
    }

    private static void setNvSources(DatabaseManager.NoiseValue nv, Set<String> sources) {
        sources = (sources == null) ? Collections.emptySet() : sources;

        nv.lift        = sources.contains("Лифт");
        nv.vent        = sources.contains("Вент");
        nv.heatCurtain = sources.contains("Завеса");
        nv.itp         = sources.contains("ИТП");
        nv.pns         = sources.contains("ПНС");
        nv.electrical  = sources.contains("Э/Щ");
        nv.autoSrc     = sources.contains("Авто");
        nv.zum         = sources.contains("Зум");
    }
    /** Применить текущий глобальный фильтр: секции → этажи → помещения → комнаты. */
    private void applyGlobalFilter() {
        commitEditors();
        // пересоздаём фильтр под текущее состояние и карту byKey
        filter = new NoiseFilter(building, byKey, getActiveFilterSources());
        refreshSections();
        refreshFloors();
        refreshSpaces();
        refreshRooms();
    }
    private int getSelectedSectionIndex() {
        Section sel = sectionList.getSelectedValue();
        if (sel == null) return 0;
        List<Section> all = building.getSections();
        int idx = all.indexOf(sel);
        return (idx < 0) ? 0 : idx;
    }

}
