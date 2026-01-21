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
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
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

    // Периоды измерений для шума (лифт день/ночь)
    private final java.util.Map<NoiseTestKind, NoisePeriod> periods = new java.util.EnumMap<>(NoiseTestKind.class);

    // четыре значения на запись: Ек мин/макс и М мин/макс
    private final Map<String, Threshold> thresholds = new LinkedHashMap<>();



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

    /** Применить сохранённые пороги (key -> {EqMin, EqMax, MaxMin, MaxMax}). */
    public void applyThresholds(java.util.Map<String, double[]> snapshot) {
        thresholds.clear();
        if (snapshot == null) return;
        snapshot.forEach((key, arr) -> thresholds.put(key, thresholdFromArray(arr)));
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

    /** Сохранить текущие пороги: key -> {EqMin, EqMax, MaxMin, MaxMax}. */
    public java.util.Map<String, double[]> saveThresholdsByKey() {
        return buildThresholdsForExport();
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
        // коммитим редактирование до смены выбора
        sectionList.addMouseListener(new MouseAdapter() {
            @Override public void mousePressed(MouseEvent e) { commitEditors(); }
        });

        floorList = new JList<>(floorModel);
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setCellRenderer(new FloorListRenderer());
        // коммитим редактирование до смены выбора
        floorList.addMouseListener(new MouseAdapter() {
            @Override public void mousePressed(MouseEvent e) { commitEditors(); }
        });

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                refreshFloors();
                refreshSpaces();
                refreshRooms();
            }
        });
        floorList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                // ВАЖНО: при смене этажа подстраиваем набор тумблеров под «улицу»
                updateSourcesCellForCurrentFloor();
                refreshSpaces();
                refreshRooms();
            }
        });

        // ==== Список помещений ====
        spaceList = new JList<>(spaceModel);
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.setCellRenderer(new SpaceListRenderer());
        // коммитим редактирование до смены выбора
        spaceList.addMouseListener(new MouseAdapter() {
            @Override public void mousePressed(MouseEvent e) { commitEditors(); }
        });
        spaceList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshRooms();
        });

        // ==== Таблица комнат (2 столбца: Комната | Источник) ====
        tableModel = new NoiseRoomsTableModel();
        roomsTable = new JTable(tableModel);
        roomsTable.setRowHeight(28);
        roomsTable.getColumnModel().getColumn(0).setPreferredWidth(220); // Комната
        roomsTable.getColumnModel().getColumn(1).setPreferredWidth(650); // Источник(тумблеры)

        // Рендер/редактор для колонки источников — с учётом «улицы»
        NoiseSourcesCell cell = new NoiseSourcesCell(getCurrentSourceLabels());
        roomsTable.getColumnModel().getColumn(1).setCellRenderer(cell);
        roomsTable.getColumnModel().getColumn(1).setCellEditor(cell);

        // На всякий: при потере фокуса таблицей — коммитим
        roomsTable.addFocusListener(new FocusAdapter() {
            @Override public void focusLost(FocusEvent e) { commitEditors(); }
        });

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

        // ==== Фильтр по источникам ====
        p.add(new JLabel("Фильтр по источникам:"));

        filterBtns = new JToggleButton[SRC_SHORT.length];
        for (int i = 0; i < SRC_SHORT.length; i++) {
            JToggleButton b = new JToggleButton(SRC_SHORT[i]);
            b.setFocusable(false);
            b.setMargin(new Insets(2, 6, 2, 6));
            b.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                    "buttonType: toolBarButton; arc: 8; focusWidth: 1");
            b.addItemListener(e -> applyGlobalFilter());
            filterBtns[i] = b;
            p.add(b);
        }

        JButton clear = new JButton("Сброс");
        clear.setFocusable(false);
        clear.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: toolBarButton; arc: 8; focusWidth: 1");
        clear.addActionListener(e -> {
            for (JToggleButton b : filterBtns) b.setSelected(false);
            applyGlobalFilter();
        });
        p.add(clear);

        // Разделитель
        JSeparator sep = new JSeparator(SwingConstants.VERTICAL);
        sep.setPreferredSize(new Dimension(10, 24));
        p.add(Box.createHorizontalStrut(4));
        p.add(sep);
        p.add(Box.createHorizontalStrut(4));

        // ==== Периоды + Экспорт ====
        JButton periodsBtn = new JButton("Периоды…");
        periodsBtn.setFocusable(false);
        periodsBtn.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: roundRect; arc: 999; minimumWidth: 110");
        periodsBtn.setToolTipText("Указать дату и время: лифт (день/ночь), ИТО (неж/жил), авто (день/ночь), площадка");
        periodsBtn.addActionListener(e -> {
            Window w = SwingUtilities.getWindowAncestor(this);
            NoisePeriodsDialog dlg = new NoisePeriodsDialog(w, periods);
            dlg.setVisible(true);
            java.util.Map<NoiseTestKind, NoisePeriod> res = dlg.getResult();
            if (res != null && !res.isEmpty()) {
                periods.clear();
                periods.putAll(res);
            }
        });
        p.add(periodsBtn);

        JButton exportBtn = new JButton("Экспорт в Excel");
        exportBtn.setFocusable(false);
        exportBtn.setIcon(org.kordamp.ikonli.swing.FontIcon.of(org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.FILE_EXCEL, 16));
        exportBtn.setIconTextGap(8);
        exportBtn.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: roundRect; background: #E6F4EA; borderColor: #34A853; arc: 999; focusWidth: 1; innerFocusWidth: 0; minimumWidth: 150");
        exportBtn.setToolTipText("Сформировать Excel по активным листам");
        exportBtn.addActionListener(e -> onExportExcel());
        p.add(exportBtn);

        // Разделитель
        JSeparator sep2 = new JSeparator(SwingConstants.VERTICAL);
        sep2.setPreferredSize(new Dimension(10, 24));
        p.add(Box.createHorizontalStrut(4));
        p.add(sep2);
        p.add(Box.createHorizontalStrut(4));

        // ==== Пороговые значения (урезанные) ====
        p.add(new JLabel("Пороговые:"));

        String[] srcForThresholds = {"Лифт","ИТО","Авто","Улица"};
        for (String src : srcForThresholds) {
            JButton b = createThresholdButton(src);
            p.add(b);
        }

        JButton exportThresholds = new JButton("Экспорт пороговых значений");
        exportThresholds.setFocusable(false);
        exportThresholds.setIcon(org.kordamp.ikonli.swing.FontIcon.of(org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.FILE_EXPORT, 16));
        exportThresholds.setIconTextGap(6);
        exportThresholds.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: toolBarButton; arc: 8; focusWidth: 1");
        exportThresholds.setToolTipText("Сформировать Excel-шаблон для пороговых значений шума");
        exportThresholds.addActionListener(e -> onExportThresholdsTemplate());
        p.add(exportThresholds);

        JButton importThresholds = new JButton("Импорт порогов");
        importThresholds.setFocusable(false);
        importThresholds.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: toolBarButton; arc: 8; focusWidth: 1");
        importThresholds.setToolTipText("Загрузить заполненный Excel-шаблон порогов");
        importThresholds.addActionListener(e -> onImportThresholdsTemplate());
        p.add(importThresholds);

        return p;
    }

    /** Экспорт «Шумы / Лифт»: создаёт все нужные листы. */
    private void onExportExcel() {
        try {
            updateRoomSelectionStates();
            Map<String, DatabaseManager.NoiseValue> snapshot = saveSelectionsByKey();

            // «Дата, время проведения измерений ...»
            java.util.EnumMap<NoiseTestKind, String> dls = new java.util.EnumMap<>(NoiseTestKind.class);
            dls.put(NoiseTestKind.LIFT_DAY,   excelDateLine(NoiseTestKind.LIFT_DAY));
            dls.put(NoiseTestKind.LIFT_NIGHT, excelDateLine(NoiseTestKind.LIFT_NIGHT));

            dls.put(NoiseTestKind.ITO_NONRES,    excelDateLine(NoiseTestKind.ITO_NONRES));
            dls.put(NoiseTestKind.ITO_RES_DAY,   excelDateLine(NoiseTestKind.ITO_RES_DAY));
            dls.put(NoiseTestKind.ITO_RES_NIGHT, excelDateLine(NoiseTestKind.ITO_RES_NIGHT));

            dls.put(NoiseTestKind.AUTO_DAY,   excelDateLine(NoiseTestKind.AUTO_DAY));
            dls.put(NoiseTestKind.AUTO_NIGHT, excelDateLine(NoiseTestKind.AUTO_NIGHT));
            dls.put(NoiseTestKind.SITE,       excelDateLine(NoiseTestKind.SITE));

            // Конвертируем внутренние Threshold -> простой Map<String,double[4]>: {EqMin,EqMax,MaxMin,MaxMax}
            Map<String, double[]> thSimple = buildThresholdsForExport();

            // Новый оверлоад экспортёра — с порогами
            NoiseExcelExporter.export(building, snapshot, this, dls, thSimple);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Ошибка экспорта: " + ex.getMessage(),
                    "Экспорт", JOptionPane.ERROR_MESSAGE);
        }
    }

    /** Построить текст "Дата, время проведения измерений ..." для нужного вида испытаний. */
    private String excelDateLine(NoiseTestKind kind) {
        NoisePeriod p = periods.get(kind);
        return (p != null) ? p.toExcelLine() : new NoisePeriod().toExcelLine();
    }

    private void onExportThresholdsTemplate() {
        try {
            Map<String, double[]> thSimple = buildThresholdsForExport();
            NoiseThresholdsExcelIO.exportTemplate(this, thSimple);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Ошибка экспорта порогов: " + ex.getMessage(),
                    "Пороговые значения", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void onImportThresholdsTemplate() {
        try {
            Map<String, double[]> imported = NoiseThresholdsExcelIO.importTemplate(this);
            if (imported == null) return;
            thresholds.clear();
            imported.forEach((key, arr) -> thresholds.put(key, thresholdFromArray(arr)));
            DatabaseManager.updateNoiseThresholds(building, imported);
            JOptionPane.showMessageDialog(this, "Пороговые значения загружены.",
                    "Пороговые значения", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Ошибка импорта порогов: " + ex.getMessage(),
                    "Пороговые значения", JOptionPane.ERROR_MESSAGE);
        }
    }
    /** Упаковка порогов в простой вид для экспортёра: key -> {EqMin, EqMax, MaxMin, MaxMax}. */
    private Map<String, double[]> buildThresholdsForExport() {
        Map<String, double[]> out = new LinkedHashMap<>();
        for (Map.Entry<String, Threshold> e : thresholds.entrySet()) {
            Threshold t = e.getValue();
            double ekMin = (t != null && t.ekMin != null) ? t.ekMin : Double.NaN;
            double ekMax = (t != null && t.ekMax != null) ? t.ekMax : Double.NaN;
            double mMin  = (t != null && t.mMin  != null) ? t.mMin  : Double.NaN;
            double mMax  = (t != null && t.mMax  != null) ? t.mMax  : Double.NaN;
            out.put(e.getKey(), new double[]{ ekMin, ekMax, mMin, mMax });
        }
        return out;
    }

    private static Threshold thresholdFromArray(double[] arr) {
        if (arr == null) return new Threshold();
        return new Threshold(
                valueFromArray(arr, 0),
                valueFromArray(arr, 1),
                valueFromArray(arr, 2),
                valueFromArray(arr, 3)
        );
    }

    private static Double valueFromArray(double[] arr, int idx) {
        if (arr == null || arr.length <= idx) return null;
        double v = arr[idx];
        return Double.isNaN(v) ? null : v;
    }

    private Set<String> getActiveFilterSources() {
        Set<String> s = new LinkedHashSet<>();
        if (filterBtns != null) {
            for (JToggleButton b : filterBtns) if (b.isSelected()) s.add(b.getText());
        }
        return s;
    }
    public void onBuildingChanged(ru.citlab24.protokol.tabs.models.Building building) {
        this.building = (building != null) ? building : new Building();
        applyGlobalFilter();
    }


    /** Применяет изменение к byKey и сразу сохраняет его в БД. */
    /** Применяет изменение к byKey и сразу сохраняет его в БД. */
    private void applyAndPersist(String key, java.util.function.Consumer<DatabaseManager.NoiseValue> change) {
        // 1) Сохраняем в оперативную карту вкладки
        DatabaseManager.NoiseValue nv = byKey.computeIfAbsent(key, k -> new DatabaseManager.NoiseValue());
        change.accept(nv);

        // 2) Точечный апсерт одной записи (устойчиво к быстрым переключениям)
        try {
            DatabaseManager.updateNoiseValueByKey(building, key, nv);
        } catch (Exception ignore) {
            // мягкая деградация: в памяти вкладки состояние всё равно уже сохранено
        }
    }



    /* ================== ВНУТРЕННЕЕ: наполнение списков ================== */

    private void refreshRooms() {
        commitEditors();

        // при каждом обновлении — убедимся, что редактор/рендер соответствуют типу этажа
        updateSourcesCellForCurrentFloor();

        Space s = spaceList.getSelectedValue();
        if (s == null) { tableModel.setRooms(Collections.emptyList()); return; }

        int secIdx = getSelectedSectionIndex();
        List<Room> base = (filter == null)
                ? s.getRooms()
                : filter.filterRooms(secIdx, floorList.getSelectedValue(), s);

        List<Room> rooms = base.stream()
                .filter(r -> !isIgnoredNoiseRoomName(r.getName()))
                .sorted(Comparator.comparingInt(Room::getPosition))
                .collect(Collectors.toList());

        tableModel.setRooms(rooms);
    }

    private void refreshFloors() {
        // ВАЖНО: сперва коммитим активный редактор, чтобы JTable успел записать значение в модель
        commitEditors();

        floorModel.clear();
        int secIdx = getSelectedSectionIndex();

        // Скрываем этажи с типом PUBLIC («общественный») в любом случае —
        // как при отсутствии, так и при наличии дополнительного фильтра NoiseFilter.
        List<Floor> floors;
        if (filter == null) {
            floors = building.getFloors().stream()
                    .filter(f -> f.getSectionIndex() == secIdx)
                    .filter(f -> f.getType() != Floor.FloorType.PUBLIC) // <- скрываем «общественный»
                    .sorted(Comparator.comparingInt(Floor::getPosition))
                    .collect(Collectors.toList());
        } else {
            floors = filter.filterFloors(secIdx).stream()
                    .filter(Objects::nonNull)
                    .filter(f -> f.getType() != Floor.FloorType.PUBLIC) // <- скрываем «общественный»
                    .sorted(Comparator.comparingInt(Floor::getPosition))
                    .collect(Collectors.toList());
        }

        floors.forEach(floorModel::addElement);
        if (!floorModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        } else {
            spaceModel.clear();
            tableModel.setRooms(Collections.emptyList());
        }
    }

    private void refreshSpaces() {
        // ВАЖНО: сперва коммитим активный редактор
        commitEditors();

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
        // Стабильный ключ на каждую строку (привязан к объекту Room текущего набора)
        private final IdentityHashMap<Room, String> stableKeyByRoom = new IdentityHashMap<>();
        private final String[] COLS = {"Комната", "Источник"};
        private List<Room> rows = new ArrayList<>();

        void setRooms(List<Room> rooms) {
            this.rows = (rooms != null) ? rooms : new ArrayList<>();
            // Перестраиваем стабильные ключи для текущего набора строк
            stableKeyByRoom.clear();

            // Фиксируем контекст прямо сейчас (он больше не будет зависеть от последующих переключений списков)
            Floor f = floorList.getSelectedValue();
            Space s = spaceList.getSelectedValue();
            int secIdx = getSelectedSectionIndex();

            String floorNum = (f == null || f.getNumber() == null) ? "" : f.getNumber().trim();
            String spaceId  = (s == null || s.getIdentifier() == null) ? "" : s.getIdentifier().trim();

            for (Room r : this.rows) {
                String roomName = (r == null || r.getName() == null) ? "" : r.getName().trim();
                String key = secIdx + "|" + floorNum + "|" + spaceId + "|" + roomName;
                stableKeyByRoom.put(r, key);
            }

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

            // Для «улицы» показываем только Авто и Поезд (Поезд ←→ nv.zum)
            if (isStreetSelected()) {
                Set<String> s2 = new LinkedHashSet<>();
                if (sources.contains("Авто")) s2.add("Авто");
                if (sources.contains("Зум"))  s2.add("Поезд");
                return s2;
            }
            return sources;
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
                Set<String> incoming = new LinkedHashSet<>((Set<String>) aValue);

                // На «улице» преобразуем «Поезд» -> «Зум», остальные отбрасываем
                Set<String> s = new LinkedHashSet<>();
                if (isStreetSelected()) {
                    if (incoming.contains("Авто"))  s.add("Авто");
                    if (incoming.contains("Поезд")) s.add("Зум"); // временно храним поезд в nv.zum
                } else {
                    s.addAll(incoming);
                }

                applyAndPersist(key, v -> {
                    setNvSources(v, s);
                    setNvSelected(v, !s.isEmpty());
                });

                if (!getActiveFilterSources().isEmpty()) {
                    SwingUtilities.invokeLater(() -> {
                        if (roomsTable != null && !roomsTable.isEditing()) {
                            refreshRooms();
                        }
                    });
                }
            }
        }

        private String keyFor(Room r) {
            // Сначала пытаемся взять стабильный ключ, зафиксированный при setRooms()
            String k = stableKeyByRoom.get(r);
            if (k != null) return k;

            // Фолбэк на случай внештатных вызовов: вычисляем по текущему выбору (как было раньше)
            Floor f = floorList.getSelectedValue();
            Space s = spaceList.getSelectedValue();
            int secIdx = getSelectedSectionIndex();

            String floorNum = (f == null || f.getNumber() == null) ? "" : f.getNumber().trim();
            String spaceId  = (s == null || s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
            String roomName = (r == null || r.getName() == null) ? "" : r.getName().trim();

            return secIdx + "|" + floorNum + "|" + spaceId + "|" + roomName;
        }
    }

    /* ================== Ячейка с набором тумблеров источников ================== */

    /* ================== Ячейка с набором тумблеров источников ================== */
    private static final class NoiseSourcesCell extends AbstractCellEditor
            implements TableCellRenderer, TableCellEditor {

        private final String[] labels;
        private JPanel panel;
        private JTable tableRef; // чтобы можно было коммитить сразу по клику

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
                // мгновенный коммит значения при каждом клике
                b.addActionListener(e -> {
                    b.setSelected(b.isSelected());
                    if (tableRef != null && tableRef.getCellEditor() != null) {
                        // зафиксировать текущее множество тумблеров
                        tableRef.getCellEditor().stopCellEditing();
                    }
                });
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
            this.tableRef = table; // сохраним ссылку для commit по клику
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
    /** Игнорируем в «Шумах» служебные помещения. */
    private static boolean isIgnoredNoiseRoomName(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(java.util.Locale.ROOT).trim();

        // базовые варианты и «как слышится»
        String[] bad = {
                "санузел", "сан узел", "сан.узел", "с/у", "с.у.", "санитарный узел",
                "совмещенный санузел", "совмещённый санузел",
                "ванная комната", "ванная",
                "коридор", "кладовая", "гардероб", "уборная"
        };
        for (String k : bad) {
            if (n.contains(k)) return true;
        }
        return false;
    }
    /** Текущий выбранный этаж — «улица»? */
    private boolean isStreetSelected() {
        Floor f = (floorList != null) ? floorList.getSelectedValue() : null;
        return f != null && f.getType() == Floor.FloorType.STREET;
    }

    /** Набор ярлыков источников для текущего этажа (обычный или «улица»). */
    private String[] getCurrentSourceLabels() {
        return isStreetSelected()
                ? new String[] { "Авто", "Поезд" }
                : SRC_SHORT;
    }

    /** Переставить редактор/рендер колонки «Источник» под текущий тип этажа. */
    private void updateSourcesCellForCurrentFloor() {
        if (roomsTable == null) return;
        NoiseSourcesCell cell = new NoiseSourcesCell(getCurrentSourceLabels());
        roomsTable.getColumnModel().getColumn(1).setCellRenderer(cell);
        roomsTable.getColumnModel().getColumn(1).setCellEditor(cell);
    }
    /** Пара «мин/макс» для одного варианта источника. */
    /** Пороговые значения: Еквивалентный и Максимальный (мин/макс). */
    private static final class Threshold {
        final Double ekMin;  // Ек мин
        final Double ekMax;  // Ек макс
        final Double mMin;   // М мин
        final Double mMax;   // М макс

        Threshold() { this(null, null, null, null); }

        Threshold(Double ekMin, Double ekMax, Double mMin, Double mMax) {
            this.ekMin = ekMin;
            this.ekMax = ekMax;
            this.mMin  = mMin;
            this.mMax  = mMax;
        }

        Threshold withEkMin(Double v) { return new Threshold(v, ekMax, mMin, mMax); }
        Threshold withEkMax(Double v) { return new Threshold(ekMin, v, mMin, mMax); }
        Threshold withMMin (Double v) { return new Threshold(ekMin, ekMax, v,    mMax); }
        Threshold withMMax (Double v) { return new Threshold(ekMin, ekMax, mMin, v   ); }
    }

    /** Варианты порогов для конкретного источника. */
    private static List<String> variantsForSource(String src) {
        if ("Лифт".equals(src) || "ИТО".equals(src)) {
            return java.util.Arrays.asList("день","ночь","офис");
        }
        if ("Авто".equals(src)) {
            return java.util.Arrays.asList("день","ночь");
        }
        if ("Улица".equals(src)) {
            return java.util.Arrays.asList("диапазон"); // одна строка
        }
        return java.util.Collections.emptyList();
    }

    private JButton createThresholdButton(String srcLabel) {
        JButton b = new JButton(srcLabel);
        b.setFocusable(false);
        b.setMargin(new Insets(2, 6, 2, 6));
        b.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: toolBarButton; arc: 8; focusWidth: 1");

        java.util.List<String> vars = variantsForSource(srcLabel);
        String hint = vars.isEmpty() ? "" : " (" + String.join("/", vars) + ")";
        b.setToolTipText("Пороговые значения: " + srcLabel + hint);

        b.addActionListener(e -> showThresholdPopup(b, srcLabel));
        return b;
    }


    /** Выпадающее меню с тремя строками: <источник> день/ночь/офис, поля «мин/макс». */
    /** Редактор порогов в виде таблицы: Ек мин/Ек макс/М мин/М макс. */
    private void showThresholdPopup(JButton owner, String srcLabel) {
        JPanel content = new JPanel(new GridBagLayout());
        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(4, 6, 4, 6);
        gc.anchor = GridBagConstraints.WEST;

        int row = 0;

        // ---- Заголовок таблицы ----
        gc.gridy = row++;
        gc.gridx = 0; gc.weightx = 1; gc.fill = GridBagConstraints.HORIZONTAL;
        JLabel title = new JLabel(srcLabel);
        title.setFont(title.getFont().deriveFont(Font.BOLD));
        content.add(title, gc);

        // строка заголовков колонок
        gc.gridy = row; gc.weightx = 0; gc.fill = GridBagConstraints.NONE;

        JLabel hVar = new JLabel("Вариант");
        gc.gridx = 0; content.add(hVar, gc);

        JLabel hEkMin = new JLabel("Eq мин");
        gc.gridx = 1; content.add(hEkMin, gc);

        JLabel hEkMax = new JLabel("Eq макс");
        gc.gridx = 2; content.add(hEkMax, gc);

        JLabel hMMin = new JLabel("MAX мин");
        gc.gridx = 3; content.add(hMMin, gc);

        JLabel hMMax = new JLabel("MAX макс");
        gc.gridx = 4; content.add(hMMax, gc);

        row++;

        // ---- Строки вариантов ----
        java.util.List<String> variants = variantsForSource(srcLabel);

        for (String variant : variants) {
            String key = thKey(srcLabel, variant);
            Threshold t = thresholds.getOrDefault(key, new Threshold());

            gc.gridy = row; gc.weightx = 0; gc.fill = GridBagConstraints.NONE;

            // Колонка 0 — подпись варианта
            JLabel lbl = new JLabel(("диапазон".equals(variant) ? "Улица" : (srcLabel + " " + variant)));
            gc.gridx = 0; content.add(lbl, gc);

            // Колонка 1 — Ек мин
            javax.swing.JFormattedTextField fEkMin = new javax.swing.JFormattedTextField(buildNumberFormatter());
            fEkMin.setColumns(6);
            fEkMin.setFocusLostBehavior(javax.swing.JFormattedTextField.PERSIST);
            if (t.ekMin != null) fEkMin.setValue(t.ekMin);
            gc.gridx = 1; content.add(fEkMin, gc);

            // Колонка 2 — Ек макс
            javax.swing.JFormattedTextField fEkMax = new javax.swing.JFormattedTextField(buildNumberFormatter());
            fEkMax.setColumns(6);
            fEkMax.setFocusLostBehavior(javax.swing.JFormattedTextField.PERSIST);
            if (t.ekMax != null) fEkMax.setValue(t.ekMax);
            gc.gridx = 2; content.add(fEkMax, gc);

            // Колонка 3 — М мин
            javax.swing.JFormattedTextField fMMin = new javax.swing.JFormattedTextField(buildNumberFormatter());
            fMMin.setColumns(6);
            fMMin.setFocusLostBehavior(javax.swing.JFormattedTextField.PERSIST);
            if (t.mMin != null) fMMin.setValue(t.mMin);
            gc.gridx = 3; content.add(fMMin, gc);

            // Колонка 4 — М макс
            javax.swing.JFormattedTextField fMMax = new javax.swing.JFormattedTextField(buildNumberFormatter());
            fMMax.setColumns(6);
            fMMax.setFocusLostBehavior(javax.swing.JFormattedTextField.PERSIST);
            if (t.mMax != null) fMMax.setValue(t.mMax);
            gc.gridx = 4; content.add(fMMax, gc);

            // сохраняем ссылки, чтобы собрать позже
            lbl.putClientProperty("key", key);
            lbl.putClientProperty("fEkMin", fEkMin);
            lbl.putClientProperty("fEkMax", fEkMax);
            lbl.putClientProperty("fMMin",  fMMin);
            lbl.putClientProperty("fMMax",  fMMax);

            row++;
        }

        // ---- Кнопки ----
        JPanel actions = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 4));
        JButton save = new JButton("Сохранить");
        JButton cancel = new JButton("Отмена");
        actions.add(save);
        actions.add(cancel);

        gc.gridy = row; gc.gridx = 0; gc.gridwidth = 5;
        gc.fill = GridBagConstraints.HORIZONTAL; gc.weightx = 1.0;
        content.add(actions, gc);

        Window window = SwingUtilities.getWindowAncestor(owner);
        JDialog dialog = new JDialog(window, "Пороговые значения: " + srcLabel, Dialog.ModalityType.APPLICATION_MODAL);
        dialog.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        dialog.setContentPane(content);
        dialog.getRootPane().setDefaultButton(save);

        // Логика кнопок
        save.addActionListener(e -> {
            for (Component c : content.getComponents()) {
                if (c instanceof JLabel lbl && lbl.getClientProperty("key") != null) {
                    String key = String.valueOf(lbl.getClientProperty("key"));
                    javax.swing.JFormattedTextField fEkMin = (javax.swing.JFormattedTextField) lbl.getClientProperty("fEkMin");
                    javax.swing.JFormattedTextField fEkMax = (javax.swing.JFormattedTextField) lbl.getClientProperty("fEkMax");
                    javax.swing.JFormattedTextField fMMin  = (javax.swing.JFormattedTextField) lbl.getClientProperty("fMMin");
                    javax.swing.JFormattedTextField fMMax  = (javax.swing.JFormattedTextField) lbl.getClientProperty("fMMax");

                    Double ekMin = fieldValueOrNull(fEkMin);
                    Double ekMax = fieldValueOrNull(fEkMax);
                    Double mMin  = fieldValueOrNull(fMMin);
                    Double mMax  = fieldValueOrNull(fMMax);

                    if (ekMin == null && ekMax == null && mMin == null && mMax == null) {
                        thresholds.remove(key);
                    } else {
                        thresholds.put(key, new Threshold(ekMin, ekMax, mMin, mMax));
                    }
                }
            }
            dialog.dispose();
        });
        cancel.addActionListener(e -> dialog.dispose());

        dialog.pack();
        dialog.setLocationRelativeTo(owner);
        dialog.setVisible(true);
    }

    private static javax.swing.text.NumberFormatter buildNumberFormatter() {
        javax.swing.text.NumberFormatter fmt = new javax.swing.text.NumberFormatter(
                java.text.NumberFormat.getNumberInstance());
        fmt.setValueClass(Double.class);
        fmt.setAllowsInvalid(true);
        return fmt;
    }

    private static Double fieldValueOrNull(javax.swing.JFormattedTextField field) {
        if (field == null) return null;
        String text = field.getText();
        if (text == null || text.trim().isEmpty()) return null;
        try {
            field.commitEdit();
        } catch (java.text.ParseException ignore) {
            return null;
        }
        Object value = field.getValue();
        return (value instanceof Number n) ? n.doubleValue() : null;
    }
    /** Ключ порога: "<источник>|<вариант>", например "Лифт|день" или "Улица|диапазон". */
    private static String thKey(String srcLabel, String variant) {
        return srcLabel + "|" + (variant == null ? "" : variant);
    }

}
