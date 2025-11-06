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
        if (building == null) return;
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (ops.isOfficeSpace(s) || ops.isPublicSpace(s)) {
                    // переносим состояния из вкладки
                    for (Room r : s.getRooms()) {
                        Boolean v = selectionMap.get(r.getId());
                        if (v != null) r.setSelected(v);
                    }
                } else {
                    // это жилищные/прочие – ВСЕГДА затираем в модели
                    for (Room r : s.getRooms()) {
                        r.setSelected(false);
                    }
                }
            }
        }
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
                    r.setSelected(selected);
                }
                // обновим подсветку этажей/помещений
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

    /**
     * Перестраиваем локальную карту.
     * Сохраняем уже выставленные значения, для новых комнат:
     *  - если помещение офис/общественное → true,
     *  - иначе (квартиры/прочее) — просто не добавляем в карту.
     */
    private void rebuildSelectionMap(boolean autoApplyDefaults) {
        if (!autoApplyDefaults) {
            // ПОЛНОСТЬЮ перечитать карту из доменной модели (сохранённые значения)
            selectionMap.clear();
            for (Floor f : building.getFloors()) {
                for (Space s : f.getSpaces()) {
                    if (!(ops.isOfficeSpace(s) || ops.isPublicSpace(s))) continue;
                    for (Room r : s.getRooms()) {
                        selectionMap.put(r.getId(), r.isSelected());
                    }
                }
            }
            return;
        }

        // autoApplyDefaults = true → не трогаем существующие ключи; для НОВЫХ комнат включаем по умолчанию
        Map<Integer, Boolean> merged = new HashMap<>(selectionMap);
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                if (!(ops.isOfficeSpace(s) || ops.isPublicSpace(s))) continue;
                for (Room r : s.getRooms()) {
                    if (!merged.containsKey(r.getId())) {
                        // если это новая комната — ставим ON по умолчанию
                        merged.put(r.getId(), true);
                    }
                }
            }
        }
        selectionMap.clear();
        selectionMap.putAll(merged);
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
        try { this.updateRoomSelectionStates(); } catch (Throwable ignore) {}

        // 3) Экспорт одного листа «Иск освещение» по всем секциям (sectionIndex = -1)
        if (this.building == null) {
            JOptionPane.showMessageDialog(this, "Сначала загрузите/выберите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        ArtificialLightingExcelExporter.export(this.building, -1, this);
    }

}
