package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.buildingTab.BuildingModelOps;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.util.*;
import java.util.List;

import org.kordamp.ikonli.swing.FontIcon;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;

import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;

/**
 * Вкладка «Освещение» (искусственное).
 * Внешне похожа на КЕО: Секции → Этажи → Помещения → Комнаты.
 * Отличия:
 *  - показываем ТОЛЬКО офисные и общественные помещения (квартиры игнорируем);
 *  - галочки по умолчанию включены для всех комнат таких помещений;
 *  - состояние хранится ЛОКАЛЬНО в этой вкладке (Map<Integer, Boolean>), модель Room не меняем.
 */
public final class ArtificialLightingTab extends JPanel {


    private Building building = new Building();
    private final BuildingModelOps ops = new BuildingModelOps(building);

    private final DefaultListModel<Section> sectionModel = new DefaultListModel<>();
    private final DefaultListModel<Floor>   floorModel   = new DefaultListModel<>();
    private final DefaultListModel<Space>   spaceModel   = new DefaultListModel<>();

    // Ярко-зелёный хайлайт для этажей/помещений с отмеченными комнатами
    private static final java.awt.Color HL_ON = new java.awt.Color(232, 245, 233); // лёгкий зелёный

    private JList<Section> sectionList;
    private JList<Floor>   floorList;
    private JList<Space>   spaceList;

    private JTable roomsTable;
    private ArtificialLightingRoomsTableModel roomsModel;

    /** Локальная (вкладочная) карта выбранности: roomId → selected. */
    private final Map<Integer, Boolean> selectionMap = new HashMap<>();

    public ArtificialLightingTab(Building building) {
        setLayout(new BorderLayout(8, 8));
        setBuilding(building, /*autoApplyDefaults=*/true);
        initUI();
        refreshData();
    }

    /** Обновление входной модели. Никаких изменений Room/Floor/Space мы не делаем. */
    public void setBuilding(Building building, boolean autoApplyDefaults) {
        this.building = (building != null) ? building : new Building();
        this.ops.setBuilding(this.building);
        rebuildSelectionMap(autoApplyDefaults);
    }

    /** Полное обновление UI-списков. */
    public void refreshData() {
        refreshSections();
        refreshFloors();
        refreshSpaces();
        refreshRooms();
    }

    /** Зафиксировать текущее состояние чекбоксов Освещения в доменной модели. */
    public void updateRoomSelectionStates() {
        // Ничего не пишем в Room — искусственное освещение живёт только в selectionMap.
        lit_refreshHighlights();
    }


    /** Выбор помещения по индексу (удобно вызывать снаружи). */
    public void selectSpaceByIndex(int index) {
        if (spaceList == null) return;
        if (index >= 0 && index < spaceModel.size()) {
            spaceList.setSelectedIndex(index);
        }
    }

    // ========================= UI =========================

    private void initUI() {
        add(buildToolbar(), BorderLayout.NORTH);
        add(buildCenter(), BorderLayout.CENTER);
        add(createExportBar(), BorderLayout.SOUTH);
    }


    private JComponent buildToolbar() {
        JToolBar tb = new JToolBar();
        tb.setFloatable(false);

        JButton btnSelectAll = new JButton("Отметить все");
        btnSelectAll.addActionListener(e -> {
            Space s = spaceList.getSelectedValue();
            if (s == null) return;
            for (Room r : s.getRooms()) selectionMap.put(r.getId(), true);
            refreshRooms();
        });

        JButton btnUnselectAll = new JButton("Снять все");
        btnUnselectAll.addActionListener(e -> {
            Space s = spaceList.getSelectedValue();
            if (s == null) return;
            for (Room r : s.getRooms()) selectionMap.put(r.getId(), false);
            refreshRooms();
        });

        tb.add(btnSelectAll);
        tb.add(btnUnselectAll);
        return tb;
    }

    private JComponent buildCenter() {
        JSplitPane split = new JSplitPane(
                JSplitPane.HORIZONTAL_SPLIT,
                buildLeftPane(),
                buildRoomsPanel()
        );
        split.setResizeWeight(0.45);
        return split;
    }
    // === Перерисовка подсветки этажей/помещений (как в КЕО) ===
    private void lit_refreshHighlights() {
        try { if (floorList != null) floorList.repaint(); } catch (Throwable ignore) {}
        try { if (spaceList != null) spaceList.repaint(); } catch (Throwable ignore) {}
    }

    private JComponent buildLeftPane() {
        // ТРИ КОЛОНКИ РЯДОМ (горизонтально), как в КЕО
        JPanel p = new JPanel(new GridLayout(1, 3, 8, 0));
        p.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 8));

        sectionList = new JList<>(sectionModel);
        floorList   = new JList<>(floorModel);
        spaceList   = new JList<>(spaceModel);

        // списокы — строго вертикальные внутри своих панелей
        sectionList.setLayoutOrientation(JList.VERTICAL);
        floorList.setLayoutOrientation(JList.VERTICAL);
        spaceList.setLayoutOrientation(JList.VERTICAL);

        sectionList.setVisibleRowCount(-1);
        floorList.setVisibleRowCount(-1);
        spaceList.setVisibleRowCount(-1);

        sectionList.setFixedCellHeight(26);
        floorList.setFixedCellHeight(26);
        spaceList.setFixedCellHeight(26);

        // Этажи: базовый рендер + ярко-зелёный фон, если на этаже есть хотя бы одна отмеченная комната
        floorList.setCellRenderer(new ListCellRenderer<Floor>() {
            private final FloorListRenderer base = new FloorListRenderer();
            @Override
            public Component getListCellRendererComponent(JList<? extends Floor> list, Floor value,
                                                          int index, boolean isSelected, boolean cellHasFocus) {
                Component c = base.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (!isSelected && c instanceof JComponent jc) {
                    boolean highlight = hasAnyOnFloor(value);
                    jc.setOpaque(true);
                    jc.setBackground(highlight ? HL_ON : UIManager.getColor("List.background"));
                }
                return c;
            }
        });

// Помещения: базовый рендер + подсветка, если в помещении есть хотя бы одна отмеченная комната
        spaceList.setCellRenderer(new ListCellRenderer<Space>() {
            private final SpaceListRenderer base = new SpaceListRenderer();
            @Override
            public Component getListCellRendererComponent(JList<? extends Space> list, Space value,
                                                          int index, boolean isSelected, boolean cellHasFocus) {
                Component c = base.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (!isSelected && c instanceof JComponent jc) {
                    boolean highlight = hasAnyInSpace(value);
                    jc.setOpaque(true);
                    jc.setBackground(highlight ? HL_ON : UIManager.getColor("List.background"));
                }
                return c;
            }
        });

        // Секциям — простой рендер по имени
        sectionList.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(
                    JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
                super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Section s) setText(s.getName() != null ? s.getName() : "");
                return this;
            }
        });

        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        sectionList.addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) refreshFloors(); });
        floorList.addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) refreshSpaces(); });
        spaceList.addListSelectionListener(e -> { if (!e.getValueIsAdjusting()) refreshRooms(); });

        // Кладём ТРИ ПАНЕЛИ РЯДОМ
        p.add(wrap("Секции", sectionList));
        p.add(wrap("Этажи", floorList));
        p.add(wrap("Помещения (офис/общественные)", spaceList));


        return p;
    }

    private JPanel buildRoomsPanel() {
        roomsModel = new ArtificialLightingRoomsTableModel(selectionMap);
        roomsTable = new JTable(roomsModel);
        roomsTable.setRowHeight(26);
        roomsTable.getColumnModel().getColumn(0).setMaxWidth(120);
        roomsTable.getColumnModel().getColumn(0).setMinWidth(110);
        roomsTable.getColumnModel().getColumn(0).setPreferredWidth(110);

// НОВОЕ: любое изменение чекбокса сразу уходит в Room.setSelected(...)
        roomsModel.addTableModelListener(e -> {
            if (e.getColumn() == 0 && e.getFirstRow() >= 0) {
                int row = e.getFirstRow();
                Room r = roomsModel.getRoomAt(row);
                if (r != null) {
                    Object v = roomsModel.getValueAt(row, 0);
                    boolean selected = (v instanceof Boolean) ? (Boolean) v : false;
                    selectionMap.put(r.getId(), selected);   // ← только локально
                }
                if (floorList != null) floorList.repaint();
                if (spaceList != null) spaceList.repaint();
            }
        });

        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(titled("Комнаты (измерения)"));
        p.add(new JScrollPane(roomsTable), BorderLayout.CENTER);
        return p;
    }

    private JPanel wrap(String title, JComponent content) {
        JPanel p = new JPanel(new BorderLayout());
        p.setBorder(titled(title));
        p.add(new JScrollPane(content), BorderLayout.CENTER);
        return p;
    }

    private TitledBorder titled(String title) {
        return BorderFactory.createTitledBorder(null, title, TitledBorder.LEFT, TitledBorder.TOP);
    }

    // ========================= DATA =========================

    private void refreshSections() {
        sectionModel.clear();
        List<Section> list = new ArrayList<>(building.getSections());
        list.sort(Comparator.comparingInt(Section::getPosition));
        list.forEach(sectionModel::addElement);
        if (!list.isEmpty() && sectionList.getSelectedIndex() < 0) sectionList.setSelectedIndex(0);
    }

    private void refreshFloors() {
        floorModel.clear();
        Section sel = sectionList.getSelectedValue();
        int secIdx = (sel == null) ? 0 : building.getSections().indexOf(sel);

        List<Floor> floors = new ArrayList<>();
        for (Floor f : building.getFloors()) {
            if (f.getSectionIndex() != secIdx) continue;
            // Скрываем ЖИЛЫЕ этажи в освещении
            if (f.getType() == Floor.FloorType.RESIDENTIAL) continue;
            floors.add(f);
        }
        floors.sort(Comparator.comparingInt(Floor::getPosition));
        floors.forEach(floorModel::addElement);

        if (!floors.isEmpty()) {
            Floor prev = floorList.getSelectedValue();
            int idx = (prev != null) ? floorModel.indexOf(prev) : -1;
            floorList.setSelectedIndex(idx >= 0 ? idx : 0);
        }
    }

    private void refreshSpaces() {
        spaceModel.clear();
        Floor f = floorList.getSelectedValue();
        if (f == null) return;

        List<Space> spaces = new ArrayList<>(f.getSpaces());
        spaces.sort(Comparator.comparingInt(Space::getPosition));
        for (Space s : spaces) {
            if (ops.isOfficeSpace(s) || ops.isPublicSpace(s)) spaceModel.addElement(s);
        }

        if (!spaceModel.isEmpty()) {
            Space prev = spaceList.getSelectedValue();
            int idx = (prev != null) ? spaceModel.indexOf(prev) : -1;
            spaceList.setSelectedIndex(idx >= 0 ? idx : 0);
        }
    }

    private void refreshRooms() {
        Space s = spaceList.getSelectedValue();
        List<Room> rooms = new ArrayList<>();
        if (s != null) rooms.addAll(s.getRooms());
        rooms.sort(Comparator.comparingInt(Room::getPosition));
        roomsModel.setRooms(rooms);
        if (floorList != null) floorList.repaint();
        if (spaceList != null) spaceList.repaint();
    }

    // ========================= SELECTION MAP =========================
    /** Снимок выбранных комнат (roomId → selected) для общего экспортёра. */
    public java.util.Map<Integer, Boolean> snapshotSelectionMap() {
        return new java.util.HashMap<>(selectionMap);
    }

/** ===== Искусственное освещение: сохранение/восстановление по ключу
 *  формат ключа совпадает с DatabaseManager.loadArtificialSelectionsByKey():
 *  sectionIndex|floor.number|space.identifier|room.name
 *  ================================================================ */

    /** Построить ключ по объектам доменной модели. */
    private static String makeKey(Floor f, Space s, Room r) {
        return f.getSectionIndex() + "|" + ns(f.getNumber()) + "|" + ns(s.getIdentifier()) + "|" + ns(r.getName());
    }
    private static String ns(String s) { return (s == null) ? "" : s.trim(); }

    /** Сохранить текущие галочки вкладки в карту по ключу.
     *  Берём только офисные и общественные помещения. */
    public java.util.Map<String, Boolean> saveSelectionsByKey() {
        java.util.Map<String, Boolean> res = new java.util.HashMap<>();
        if (building == null) return res;

        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!(ops.isOfficeSpace(s) || ops.isPublicSpace(s))) continue;
                for (Room r : s.getRooms()) {
                    boolean sel = selectionMap.getOrDefault(r.getId(), false);
                    res.put(makeKey(f, s, r), sel);
                }
            }
        }
        return res;
    }

    /** Применить карту галочек по ключу к ТЕКУЩЕЙ вкладке.
     *  Ничего не пишет в Room — заполняет только selectionMap. */
    public void applySelectionsByKey(Building b, java.util.Map<String, Boolean> byKey) {
        if (b == null || byKey == null) return;

        // обновляем ссылку на здание и ops
        this.building = b;
        this.ops.setBuilding(this.building);

        // перестраиваем selectionMap под актуальные roomId, без авто-дефолтов
        rebuildSelectionMap(/*autoApplyDefaults=*/false);

        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!(ops.isOfficeSpace(s) || ops.isPublicSpace(s))) continue;
                for (Room r : s.getRooms()) {
                    String key = makeKey(f, s, r);
                    Boolean v = byKey.get(key);
                    if (v != null) selectionMap.put(r.getId(), v);
                }
            }
        }
        // Подсветим списки (UI перерисуешь снаружи через refreshData())
        lit_refreshHighlights();
    }

    private void rebuildSelectionMap(boolean autoApplyDefaults) {
        if (building == null) {
            selectionMap.clear();
            return;
        }

        // 1) Какие roomId сейчас «валидны» (только офис/общественные)
        Set<Integer> validIds = new HashSet<>();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!(ops.isOfficeSpace(s) || ops.isPublicSpace(s))) continue;
                for (Room r : s.getRooms()) validIds.add(r.getId());
            }
        }

        // 2) Сохраняем старые значения
        Map<Integer, Boolean> prev = new HashMap<>(selectionMap);
        selectionMap.clear();

        // 3) Переносим, что знаем; для новых — по умолчанию true (если autoApplyDefaults)
        for (Integer id : validIds) {
            if (prev.containsKey(id)) {
                selectionMap.put(id, prev.get(id));
            } else if (autoApplyDefaults) {
                selectionMap.put(id, true);
            }
            // иначе — не добавляем (false по умолчанию)
        }
    }


    /** Есть ли в помещении хотя бы одна комната, отмеченная для освещения? */
    private boolean hasAnyInSpace(Space s) {
        if (s == null) return false;
        for (Room r : s.getRooms()) {
            if (selectionMap.getOrDefault(r.getId(), false)) return true;
        }
        return false;
    }

    /** Есть ли на этаже хотя бы одно помещение с отмеченными комнатами? */
    private boolean hasAnyOnFloor(Floor f) {
        if (f == null) return false;
        for (Space s : f.getSpaces()) {
            if (hasAnyInSpace(s)) return true;
        }
        return false;
    }

    /** Нижняя панель с кнопкой «Экспорт в Excel (искусств.)» */
    private JComponent createExportBar() {
        JPanel bar = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 6));
        JButton btn = new JButton("Экспорт: искусственное освещение", FontIcon.of(FontAwesomeSolid.FILE_EXCEL, 16, Color.WHITE));
        btn.setBackground(new Color(106, 27, 154)); // фиолетовый, как договорились
        btn.setForeground(Color.WHITE);
        btn.setFocusPainted(false);
        btn.addActionListener(this::onExportExcel);
        bar.add(btn);
        return bar;
    }

    /** Обработчик экспорта */
    private void onExportExcel(ActionEvent e) {
        // 1) Зафиксируем активное редактирование таблиц, если открыто
        try {
            KeyboardFocusManager kfm = KeyboardFocusManager.getCurrentKeyboardFocusManager();
            Component fo = (kfm != null) ? kfm.getFocusOwner() : null;
            JTable editingTable = (fo == null) ? null
                    : (JTable) SwingUtilities.getAncestorOfClass(JTable.class, fo);
            if (editingTable != null && editingTable.isEditing()) {
                try { editingTable.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        } catch (Exception ignore) {}

        // 2) Синхронизируем чекбоксы вкладки в модель комнат (если у вас есть такой метод)
        Map<Integer, Boolean> snap = new HashMap<>(selectionMap);

        // 3) Экспорт одного листа «Иск освещение» по всем секциям (sectionIndex = -1)
        if (this.building == null) {
            JOptionPane.showMessageDialog(this, "Сначала загрузите/выберите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        ArtificialLightingExcelExporter.export(this.building, -1, this, snap);
    }

}
