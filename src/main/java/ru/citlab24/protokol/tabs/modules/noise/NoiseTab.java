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
    private JDialog noiseSummaryDialog;
    private JTable noiseSummaryTable;
    private NoiseSummaryTableModel noiseSummaryModel;
    private JTextField noiseSummaryFilterField;
    private javax.swing.table.TableRowSorter<NoiseSummaryTableModel> noiseSummarySorter;

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
        refreshNoiseSummaryWindow();
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

        JButton summaryBtn = new JButton("Сводка точек");
        summaryBtn.setFocusable(false);
        summaryBtn.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: toolBarButton; arc: 8; focusWidth: 1");
        summaryBtn.setToolTipText("Показать выбранные шумовые точки и источники звука");
        summaryBtn.addActionListener(e -> showNoiseSummaryWindow());
        p.add(summaryBtn);

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
            updateRoomSelectionStates();
            java.util.Map<NoiseTestKind, Integer> points = buildPointsCountByKind();
            NoisePeriodsDialog dlg = new NoisePeriodsDialog(w, periods, points);
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

        JButton thresholdsMenu = new JButton("Пороговые значения");
        thresholdsMenu.setFocusable(false);
        thresholdsMenu.setIcon(org.kordamp.ikonli.swing.FontIcon.of(org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.SLIDERS_H, 16));
        thresholdsMenu.setIconTextGap(6);
        thresholdsMenu.putClientProperty(com.formdev.flatlaf.FlatClientProperties.STYLE,
                "buttonType: roundRect; arc: 999; focusWidth: 1; minimumWidth: 165");
        thresholdsMenu.setToolTipText("Настроить, экспортировать или импортировать пороговые значения шума");
        thresholdsMenu.addActionListener(e -> showThresholdsDialog(thresholdsMenu));
        p.add(thresholdsMenu);

        return p;
    }

    private void showThresholdsDialog(JButton owner) {
        String[] srcForThresholds = {"Лифт", "ИТО", "Зум", "Авто", "Улица"};

        JPanel table = new JPanel(new GridBagLayout());
        table.setBorder(BorderFactory.createEmptyBorder(6, 8, 6, 8));
        java.util.List<JLabel> editorLabels = new ArrayList<>();

        GridBagConstraints gc = new GridBagConstraints();
        gc.insets = new Insets(4, 6, 4, 6);
        gc.anchor = GridBagConstraints.WEST;

        int row = 0;
        Font headerFont = table.getFont().deriveFont(Font.BOLD);

        gc.gridy = row++;
        gc.gridx = 0; gc.weightx = 1.0; gc.fill = GridBagConstraints.HORIZONTAL;
        JLabel hVar = new JLabel("Источник / вариант");
        hVar.setFont(headerFont);
        table.add(hVar, gc);

        gc.weightx = 0; gc.fill = GridBagConstraints.NONE;
        JLabel hEkMin = new JLabel("Eq мин");
        hEkMin.setFont(headerFont);
        gc.gridx = 1; table.add(hEkMin, gc);

        JLabel hEkMax = new JLabel("Eq макс");
        hEkMax.setFont(headerFont);
        gc.gridx = 2; table.add(hEkMax, gc);

        JLabel hMMin = new JLabel("MAX мин");
        hMMin.setFont(headerFont);
        gc.gridx = 3; table.add(hMMin, gc);

        JLabel hMMax = new JLabel("MAX макс");
        hMMax.setFont(headerFont);
        gc.gridx = 4; table.add(hMMax, gc);

        for (String src : srcForThresholds) {
            gc.gridy = row++;
            gc.gridx = 0;
            gc.gridwidth = 5;
            gc.weightx = 1.0;
            gc.fill = GridBagConstraints.HORIZONTAL;
            gc.insets = new Insets(row == 2 ? 8 : 14, 6, 2, 6);
            JLabel section = new JLabel(src);
            section.setFont(section.getFont().deriveFont(Font.BOLD));
            table.add(section, gc);

            gc.gridwidth = 1;
            gc.insets = new Insets(4, 6, 4, 6);

            for (String variant : variantsForSource(src)) {
                row = addThresholdEditorRow(table, gc, row, editorLabels, src, variant);
            }
        }

        gc.gridy = row;
        gc.gridx = 0;
        gc.gridwidth = 5;
        gc.weightx = 1.0;
        gc.weighty = 1.0;
        gc.fill = GridBagConstraints.BOTH;
        table.add(Box.createGlue(), gc);

        JScrollPane scroll = new JScrollPane(table);
        scroll.setBorder(BorderFactory.createEmptyBorder());
        scroll.getVerticalScrollBar().setUnitIncrement(16);

        JPanel content = new JPanel(new BorderLayout(0, 8));
        content.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));
        content.add(scroll, BorderLayout.CENTER);

        JButton exportBtn = new JButton("Экспорт");
        exportBtn.setIcon(org.kordamp.ikonli.swing.FontIcon.of(org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.FILE_EXPORT, 14));
        JButton importBtn = new JButton("Импорт");
        importBtn.setIcon(org.kordamp.ikonli.swing.FontIcon.of(org.kordamp.ikonli.fontawesome5.FontAwesomeSolid.FILE_IMPORT, 14));
        JPanel leftActions = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        leftActions.add(importBtn);
        leftActions.add(exportBtn);

        JButton save = new JButton("Сохранить");
        JButton cancel = new JButton("Отмена");
        JPanel rightActions = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 0));
        rightActions.add(save);
        rightActions.add(cancel);

        JPanel actions = new JPanel(new BorderLayout());
        actions.add(leftActions, BorderLayout.WEST);
        actions.add(rightActions, BorderLayout.EAST);
        content.add(actions, BorderLayout.SOUTH);

        Window window = SwingUtilities.getWindowAncestor(owner);
        JDialog dialog = new JDialog(window, "Пороговые значения", Dialog.ModalityType.APPLICATION_MODAL);
        dialog.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);
        dialog.setContentPane(content);
        dialog.getRootPane().setDefaultButton(save);

        save.addActionListener(e -> {
            if (saveThresholdEditors(editorLabels, dialog)) {
                dialog.dispose();
            }
        });
        cancel.addActionListener(e -> dialog.dispose());
        exportBtn.addActionListener(e -> {
            if (saveThresholdEditors(editorLabels, dialog)) {
                onExportThresholdsTemplate();
            }
        });
        importBtn.addActionListener(e -> {
            if (onImportThresholdsTemplate()) {
                dialog.dispose();
                SwingUtilities.invokeLater(() -> showThresholdsDialog(owner));
            }
        });

        dialog.setPreferredSize(new Dimension(720, 780));
        dialog.pack();
        dialog.setLocationRelativeTo(owner);
        dialog.setVisible(true);
    }

    private int addThresholdEditorRow(JPanel content, GridBagConstraints gc, int row,
                                      java.util.List<JLabel> editorLabels, String srcLabel, String variant) {
        String key = thKey(srcLabel, variant);
        Threshold t = thresholds.getOrDefault(key, new Threshold());

        gc.gridy = row;
        gc.gridwidth = 1;
        gc.weighty = 0;
        gc.fill = GridBagConstraints.NONE;

        JLabel lbl = new JLabel("диапазон".equals(variant) ? "Улица" : (srcLabel + " " + variant));
        gc.gridx = 0;
        gc.weightx = 1.0;
        gc.fill = GridBagConstraints.HORIZONTAL;
        content.add(lbl, gc);

        javax.swing.JFormattedTextField fEkMin = createThresholdField(t.ekMin);
        gc.gridx = 1;
        gc.weightx = 0;
        gc.fill = GridBagConstraints.NONE;
        content.add(fEkMin, gc);

        javax.swing.JFormattedTextField fEkMax = createThresholdField(t.ekMax);
        gc.gridx = 2;
        content.add(fEkMax, gc);

        javax.swing.JFormattedTextField fMMin = createThresholdField(t.mMin);
        gc.gridx = 3;
        content.add(fMMin, gc);

        javax.swing.JFormattedTextField fMMax = createThresholdField(t.mMax);
        gc.gridx = 4;
        content.add(fMMax, gc);

        lbl.putClientProperty("key", key);
        lbl.putClientProperty("fEkMin", fEkMin);
        lbl.putClientProperty("fEkMax", fEkMax);
        lbl.putClientProperty("fMMin",  fMMin);
        lbl.putClientProperty("fMMax",  fMMax);
        editorLabels.add(lbl);

        return row + 1;
    }

    private static javax.swing.JFormattedTextField createThresholdField(Double value) {
        javax.swing.JFormattedTextField field = new javax.swing.JFormattedTextField(buildNumberFormatter());
        field.setColumns(6);
        field.setFocusLostBehavior(javax.swing.JFormattedTextField.PERSIST);
        if (value != null) {
            field.setValue(value);
        }
        return field;
    }

    private boolean saveThresholdEditors(java.util.List<JLabel> editorLabels, Component parent) {
        Map<String, Threshold> updates = new LinkedHashMap<>();
        Set<String> removals = new LinkedHashSet<>();

        for (JLabel lbl : editorLabels) {
            String key = String.valueOf(lbl.getClientProperty("key"));
            javax.swing.JFormattedTextField fEkMin = (javax.swing.JFormattedTextField) lbl.getClientProperty("fEkMin");
            javax.swing.JFormattedTextField fEkMax = (javax.swing.JFormattedTextField) lbl.getClientProperty("fEkMax");
            javax.swing.JFormattedTextField fMMin  = (javax.swing.JFormattedTextField) lbl.getClientProperty("fMMin");
            javax.swing.JFormattedTextField fMMax  = (javax.swing.JFormattedTextField) lbl.getClientProperty("fMMax");

            Double ekMin = fieldValueOrNull(fEkMin);
            Double ekMax = fieldValueOrNull(fEkMax);
            Double mMin  = fieldValueOrNull(fMMin);
            Double mMax  = fieldValueOrNull(fMMax);

            if (!validateRange(lbl.getText(), "Eq", ekMin, ekMax, parent)) {
                return false;
            }
            if (!validateRange(lbl.getText(), "MAX", mMin, mMax, parent)) {
                return false;
            }

            if (ekMin == null && ekMax == null && mMin == null && mMax == null) {
                removals.add(key);
            } else {
                updates.put(key, new Threshold(ekMin, ekMax, mMin, mMax));
            }
        }

        for (String key : removals) {
            thresholds.remove(key);
        }
        thresholds.putAll(updates);
        return true;
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
        if (p == null && kind == NoiseTestKind.ITO_RES_DAY) {
            p = periods.get(NoiseTestKind.ZUM_DAY);
        }
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

    private boolean onImportThresholdsTemplate() {
        try {
            Map<String, double[]> imported = NoiseThresholdsExcelIO.importTemplate(this);
            if (imported == null) return false;
            thresholds.clear();
            imported.forEach((key, arr) -> thresholds.put(key, thresholdFromArray(arr)));
            DatabaseManager.updateNoiseThresholds(building, imported);
            JOptionPane.showMessageDialog(this, "Пороговые значения загружены.",
                    "Пороговые значения", JOptionPane.INFORMATION_MESSAGE);
            return true;
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Ошибка импорта порогов: " + ex.getMessage(),
                    "Пороговые значения", JOptionPane.ERROR_MESSAGE);
            return false;
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
        refreshNoiseSummaryWindow();
    }


    /** Применяет изменение к byKey и сразу сохраняет его в БД. */
    /** Применяет изменение к byKey и сразу сохраняет его в БД. */
    private void applyAndPersist(String key, java.util.function.Consumer<DatabaseManager.NoiseValue> change) {
        // 1) Сохраняем в оперативную карту вкладки
        DatabaseManager.NoiseValue nv = byKey.computeIfAbsent(key, k -> new DatabaseManager.NoiseValue());
        change.accept(nv);
        refreshNoiseSummaryWindow();

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
        refreshNoiseSummaryWindow();
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

    private void showNoiseSummaryWindow() {
        commitEditors();
        if (noiseSummaryDialog == null || !noiseSummaryDialog.isDisplayable()) {
            Window owner = SwingUtilities.getWindowAncestor(this);
            noiseSummaryDialog = new JDialog(owner, "Сводка шумовых точек", Dialog.ModalityType.MODELESS);
            noiseSummaryDialog.setDefaultCloseOperation(WindowConstants.DISPOSE_ON_CLOSE);

            noiseSummaryModel = new NoiseSummaryTableModel();
            noiseSummaryTable = new JTable(noiseSummaryModel);
            noiseSummaryTable.setRowHeight(26);
            noiseSummarySorter = new javax.swing.table.TableRowSorter<>(noiseSummaryModel);
            noiseSummaryTable.setRowSorter(noiseSummarySorter);
            noiseSummaryTable.setDefaultRenderer(Object.class, new NoiseSummaryCellRenderer());
            noiseSummaryTable.getColumnModel().getColumn(0).setPreferredWidth(90);
            noiseSummaryTable.getColumnModel().getColumn(1).setPreferredWidth(320);
            noiseSummaryTable.getColumnModel().getColumn(2).setPreferredWidth(260);

            noiseSummaryFilterField = new JTextField();
            noiseSummaryFilterField.putClientProperty(com.formdev.flatlaf.FlatClientProperties.PLACEHOLDER_TEXT,
                    "Фильтр по этажу, комнате или источнику");
            noiseSummaryFilterField.getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
                @Override public void insertUpdate(javax.swing.event.DocumentEvent e) { applyNoiseSummaryFilter(); }
                @Override public void removeUpdate(javax.swing.event.DocumentEvent e) { applyNoiseSummaryFilter(); }
                @Override public void changedUpdate(javax.swing.event.DocumentEvent e) { applyNoiseSummaryFilter(); }
            });

            JPanel filterPanel = new JPanel(new BorderLayout(8, 0));
            filterPanel.add(new JLabel("Фильтр:"), BorderLayout.WEST);
            filterPanel.add(noiseSummaryFilterField, BorderLayout.CENTER);

            JPanel content = new JPanel(new BorderLayout(0, 6));
            content.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
            content.add(filterPanel, BorderLayout.NORTH);
            JScrollPane tableScroll = new JScrollPane(noiseSummaryTable);
            tableScroll.setVerticalScrollBarPolicy(ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
            tableScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
            tableScroll.getVerticalScrollBar().setUnitIncrement(16);
            content.add(tableScroll, BorderLayout.CENTER);

            noiseSummaryDialog.setContentPane(content);
            noiseSummaryDialog.setSize(780, 460);
            noiseSummaryDialog.setLocationRelativeTo(this);
        }

        refreshNoiseSummaryWindow();
        noiseSummaryDialog.setVisible(true);
        noiseSummaryDialog.toFront();
    }

    private void refreshNoiseSummaryWindow() {
        if (noiseSummaryModel != null) {
            noiseSummaryModel.setRows(buildNoiseSummaryRows());
            applyNoiseSummaryFilter();
        }
    }

    private void applyNoiseSummaryFilter() {
        if (noiseSummarySorter == null || noiseSummaryFilterField == null) return;

        String query = safeTrim(noiseSummaryFilterField.getText()).toLowerCase(Locale.ROOT);
        if (query.isBlank()) {
            noiseSummarySorter.setRowFilter(null);
            return;
        }

        noiseSummarySorter.setRowFilter(new RowFilter<NoiseSummaryTableModel, Integer>() {
            @Override
            public boolean include(Entry<? extends NoiseSummaryTableModel, ? extends Integer> entry) {
                for (int col = 0; col < entry.getValueCount(); col++) {
                    String value = safeTrim(entry.getStringValue(col)).toLowerCase(Locale.ROOT);
                    if (value.contains(query)) {
                        return true;
                    }
                }
                return false;
            }
        });
    }

    private List<NoiseSummaryRow> buildNoiseSummaryRows() {
        List<NoiseSummaryRow> result = new ArrayList<>();
        if (building == null) return result;

        List<Floor> floors = new ArrayList<>(building.getFloors());
        floors.sort(Comparator.comparingInt(Floor::getPosition));

        Map<String, Integer> colorBySpace = new LinkedHashMap<>();
        int colorIndex = 0;

        for (Floor floor : floors) {
            if (floor == null || floor.getType() == Floor.FloorType.PUBLIC) continue;

            String floorNum = safeTrim(floor.getNumber());
            String floorText = floorNum.isBlank() ? safeTrim(floor.getName()) : floorNum;
            int sectionIndex = Math.max(0, floor.getSectionIndex());

            List<Space> spaces = new ArrayList<>(floor.getSpaces());
            spaces.sort(Comparator.comparingInt(Space::getPosition));
            for (Space space : spaces) {
                if (space == null) continue;

                String spaceId = safeTrim(space.getIdentifier());
                String spaceLabel = spaceId.isBlank() ? spaceDisplayName(space) : spaceId;
                String spaceKey = sectionIndex + "|" + floorNum + "|" + spaceId;
                Integer existingColor = colorBySpace.get(spaceKey);
                if (existingColor == null) {
                    existingColor = colorIndex++;
                    colorBySpace.put(spaceKey, existingColor);
                }

                List<Room> rooms = new ArrayList<>(space.getRooms());
                rooms.sort(Comparator.comparingInt(Room::getPosition));
                for (Room room : rooms) {
                    if (room == null || isIgnoredNoiseRoomName(room.getName())) continue;

                    String roomName = safeTrim(room.getName());
                    String key = sectionIndex + "|" + floorNum + "|" + spaceId + "|" + roomName;
                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    Set<String> sources = (nv == null) ? Collections.emptySet() : getNvSources(nv);
                    if (sources.isEmpty()) continue;

                    result.add(new NoiseSummaryRow(
                            floorText,
                            joinComma(spaceLabel, roomName),
                            formatSourcesForSummary(floor, space, sources),
                            existingColor
                    ));
                }
            }
        }

        return result;
    }

    private static String formatSourcesForSummary(Floor floor, Space space, Set<String> sources) {
        List<String> labels = new ArrayList<>();
        boolean isStreet = floor != null && floor.getType() == Floor.FloorType.STREET
                && space != null && space.getType() == Space.SpaceType.OUTDOOR;
        for (String source : sources) {
            if (isStreet && "Зум".equals(source)) {
                labels.add("Поезд");
            } else {
                labels.add(source);
            }
        }
        return String.join(", ", labels);
    }

    private static String joinComma(String... values) {
        List<String> parts = new ArrayList<>();
        if (values != null) {
            for (String value : values) {
                String trimmed = safeTrim(value);
                if (!trimmed.isBlank()) parts.add(trimmed);
            }
        }
        return String.join(", ", parts);
    }

    private static String safeTrim(String value) {
        return value == null ? "" : value.trim();
    }

    private static String spaceDisplayName(Space space) {
        if (space == null) return "";
        String identifier = safeTrim(space.getIdentifier());
        if (!identifier.isBlank()) return identifier;
        try {
            java.lang.reflect.Method m = space.getClass().getMethod("getName");
            Object value = m.invoke(space);
            if (value != null && !value.toString().trim().isBlank()) {
                return value.toString().trim();
            }
        } catch (Exception ignored) {
        }
        return "Помещение";
    }

    private static final class NoiseSummaryRow {
        final String floor;
        final String room;
        final String sources;
        final int colorIndex;

        NoiseSummaryRow(String floor, String room, String sources, int colorIndex) {
            this.floor = floor;
            this.room = room;
            this.sources = sources;
            this.colorIndex = colorIndex;
        }
    }

    private static final class NoiseSummaryTableModel extends AbstractTableModel {
        private final String[] columns = {"Этаж", "Комната в помещении", "Источник звука"};
        private List<NoiseSummaryRow> rows = new ArrayList<>();

        void setRows(List<NoiseSummaryRow> rows) {
            this.rows = rows == null ? new ArrayList<>() : new ArrayList<>(rows);
            fireTableDataChanged();
        }

        NoiseSummaryRow rowAt(int rowIndex) {
            return rows.get(rowIndex);
        }

        @Override public int getRowCount() { return rows.size(); }
        @Override public int getColumnCount() { return columns.length; }
        @Override public String getColumnName(int column) { return columns[column]; }

        @Override public Object getValueAt(int rowIndex, int columnIndex) {
            NoiseSummaryRow row = rows.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> row.floor;
                case 1 -> row.room;
                case 2 -> row.sources;
                default -> "";
            };
        }
    }

    private static final class NoiseSummaryCellRenderer extends javax.swing.table.DefaultTableCellRenderer {
        private static final Color[] COLORS = {
                new Color(232, 244, 255),
                new Color(236, 248, 239),
                new Color(255, 245, 225),
                new Color(244, 238, 255),
                new Color(255, 236, 239),
                new Color(232, 248, 247)
        };

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                                                       boolean hasFocus, int row, int column) {
            Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
            if (!isSelected && table.getModel() instanceof NoiseSummaryTableModel model) {
                int modelRow = table.convertRowIndexToModel(row);
                NoiseSummaryRow summaryRow = model.rowAt(modelRow);
                c.setBackground(COLORS[Math.floorMod(summaryRow.colorIndex, COLORS.length)]);
            }
            return c;
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
        refreshNoiseSummaryWindow();
    }
    private int getSelectedSectionIndex() {
        Section sel = sectionList.getSelectedValue();
        if (sel == null) return 0;
        List<Section> all = building.getSections();
        int idx = all.indexOf(sel);
        return (idx < 0) ? 0 : idx;
    }
    private java.util.Map<NoiseTestKind, Integer> buildPointsCountByKind() {
        java.util.EnumMap<NoiseTestKind, Integer> counts = new java.util.EnumMap<>(NoiseTestKind.class);
        for (NoiseTestKind kind : NoiseTestKind.values()) {
            counts.put(kind, 0);
        }

        if (building == null) return counts;

        java.util.List<Floor> floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor fl : floors) {
            String floorNum = (fl.getNumber() == null) ? "" : fl.getNumber().trim();
            java.util.List<Space> spaces = new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space sp : spaces) {
                String spaceId = (sp.getIdentifier() == null) ? "" : sp.getIdentifier().trim();
                java.util.List<Room> rooms = new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room rm : rooms) {
                    String roomName = (rm.getName() == null) ? "" : rm.getName().trim();
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;
                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null) continue;

                    if (nv.lift) {
                        if (sp.getType() == Space.SpaceType.APARTMENT) {
                            addPoints(counts, NoiseTestKind.LIFT_DAY, 3);
                            addPoints(counts, NoiseTestKind.LIFT_NIGHT, 3);
                        } else if (sp.getType() == Space.SpaceType.OFFICE) {
                            addPoints(counts, NoiseTestKind.LIFT_DAY, 3);
                        }
                    }

                    if (hasItoSources(nv)) {
                        if (sp.getType() == Space.SpaceType.OFFICE || sp.getType() == Space.SpaceType.PUBLIC_SPACE) {
                            addPoints(counts, NoiseTestKind.ITO_NONRES, 3);
                        }
                        if (sp.getType() == Space.SpaceType.APARTMENT) {
                            addPoints(counts, NoiseTestKind.ITO_RES_DAY, 3);
                            addPoints(counts, NoiseTestKind.ITO_RES_NIGHT, 3);
                        }
                    }

                    if (nv.zum && sp.getType() == Space.SpaceType.APARTMENT) {
                        addPoints(counts, NoiseTestKind.ITO_RES_DAY, 3);
                    }

                    if (nv.autoSrc && sp.getType() == Space.SpaceType.APARTMENT) {
                        addPoints(counts, NoiseTestKind.AUTO_DAY, 3);
                        addPoints(counts, NoiseTestKind.AUTO_NIGHT, 3);
                    }

                    if (fl.getType() == Floor.FloorType.STREET
                            && sp.getType() == Space.SpaceType.OUTDOOR
                            && (nv.autoSrc || nv.zum)) {
                        addPoints(counts, NoiseTestKind.SITE, 3);
                    }
                }
            }
        }

        return counts;
    }

    private static void addPoints(java.util.Map<NoiseTestKind, Integer> counts, NoiseTestKind kind, int delta) {
        if (counts == null || kind == null) return;
        counts.put(kind, counts.getOrDefault(kind, 0) + delta);
    }

    private static boolean hasItoSources(DatabaseManager.NoiseValue nv) {
        return nv != null && (nv.vent || nv.heatCurtain || nv.itp || nv.pns || nv.electrical);
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
        if ("Зум".equals(src)) {
            return java.util.Arrays.asList("день");
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

                    if (!validateRange(lbl.getText(), "Eq", ekMin, ekMax, dialog)) {
                        return;
                    }
                    if (!validateRange(lbl.getText(), "MAX", mMin, mMax, dialog)) {
                        return;
                    }

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

    private static boolean validateRange(String label, String kind, Double min, Double max, Component parent) {
        if (min == null || max == null) return true;
        double range = Math.abs(max - min);
        if (range < 0.6) {
            JOptionPane.showMessageDialog(parent,
                    "Диапазон для \"" + label + "\" (" + kind + ") должен быть не меньше 0,6.",
                    "Пороговые значения", JOptionPane.WARNING_MESSAGE);
            return false;
        }
        return true;
    }
    /** Ключ порога: "<источник>|<вариант>", например "Лифт|день" или "Улица|диапазон". */
    private static String thKey(String srcLabel, String variant) {
        return srcLabel + "|" + (variant == null ? "" : variant);
    }

}
