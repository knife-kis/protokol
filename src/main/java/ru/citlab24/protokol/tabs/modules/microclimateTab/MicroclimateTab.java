package ru.citlab24.protokol.tabs.modules.microclimateTab;

import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.renderers.FloorListRenderer;
import ru.citlab24.protokol.tabs.renderers.SpaceListRenderer;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionListener;
import java.util.List;
import java.util.*;
import java.util.stream.Collectors;

public class MicroclimateTab extends JPanel {

    private static final Font HEADER_FONT =
            UIManager.getFont("Label.font").deriveFont(Font.PLAIN, 15f);
    // ===== Подсветка (микроклимат) ===============================================
    private static final java.awt.Color HL_GREEN = new java.awt.Color(232, 245, 233); // лёгкий зелёный

    // Порядок: как в модели (по position)
    private static final Comparator<Section> SECTION_ORDER =
            Comparator.comparingInt(Section::getPosition);
    private static final Comparator<Floor> FLOOR_ORDER =
            Comparator.comparingInt(Floor::getPosition);
    private static final Comparator<Space> SPACE_ORDER =
            Comparator.comparingInt(Space::getPosition);
    private static final Comparator<Room> ROOM_ORDER =
            Comparator.comparingInt(Room::getPosition);

    // Модель здания
    private Building currentBuilding;

    // Секции/Этажи
    private final DefaultListModel<Section> sectionListModel = new DefaultListModel<>();
    private final JList<Section> sectionList = new JList<>(sectionListModel);

    private final DefaultListModel<Floor> floorListModel = new DefaultListModel<>();
    private final JList<Floor> floorList = new JList<>(floorListModel);

    // Помещения (таблица)
    private final SpaceTableModel spaceTableModel = new SpaceTableModel();
    private final JTable spaceTable = new JTable(spaceTableModel);

    // Комнаты (таблица) — чекбокс + название + «Наружные стены (0..4)»
    public final Map<Integer, Boolean> globalRoomSelectionMap = new HashMap<>();
    private final MicroclimateRoomsTableModel roomTableModel =
            new MicroclimateRoomsTableModel(globalRoomSelectionMap);
    private final JTable roomTable = new JTable(roomTableModel);
    private MicroclimateExcelExporter.TemperatureMode temperatureMode =
            MicroclimateExcelExporter.TemperatureMode.COLD;

    public MicroclimateTab() {
        setLayout(new BorderLayout(10, 0));
        add(buildLeftPanel(), BorderLayout.WEST);
        add(buildCenterPanel(), BorderLayout.CENTER);
        add(buildRightPanel(), BorderLayout.EAST);
        add(buildBottomPanel(), BorderLayout.SOUTH);
    }

    // ==== Публичные методы, которые вызывает BuildingTab ====

    public void display(Building building, boolean autoApplyDefaults) {
        this.currentBuilding = building;

        globalRoomSelectionMap.clear();
        for (Floor f : building.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    globalRoomSelectionMap.put(r.getId(), r.isMicroclimateSelected());
                }
            }
        }
        refreshSections();
        refreshFloors();
        micro_refreshHighlights();
    }

    /** Перелить состояния чекбоксов обратно в доменную модель (перед сохранением проекта). */
    // ==== Публичные методы, которые вызывает BuildingTab ====
    public void updateRoomSelectionStates() {
        if (currentBuilding == null) return;
        for (Floor f : currentBuilding.getFloors()) {
            for (Space s : f.getSpaces()) {
                for (Room r : s.getRooms()) {
                    Boolean val = globalRoomSelectionMap.get(r.getId());
                    r.setMicroclimateSelected(val != null && val);
                }
            }
        }
        micro_refreshHighlights();
    }

    // ==== UI: левые панели ====

    private JPanel buildLeftPanel() {
        JPanel left = new JPanel(new BorderLayout(0, 10));

        // Секции
        JPanel secPanel = new JPanel(new BorderLayout());
        secPanel.setBorder(new TitledBorder("Секции"));
        sectionList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        sectionList.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(
                    JList<?> list, Object value, int index,
                    boolean isSelected, boolean cellHasFocus) {
                super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Section s) setText(s.getName());
                else setText(value != null ? value.toString() : "");
                return this;
            }
        });

        sectionList.addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) refreshFloors();
        });
        secPanel.add(new JScrollPane(sectionList), BorderLayout.CENTER);

        // Этажи + кнопка «Проставить галочки на этаже (жилые)»
        JPanel floorPanel = new JPanel(new BorderLayout());
        floorPanel.setBorder(new TitledBorder("Этажи"));
        floorList.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);

        // БЫЛО:
        // floorList.setCellRenderer(new FloorListRenderer());

        // СТАЛО: рендерер с подсветкой этажей, где есть хотя бы одна галочка микроклимата
        floorList.setCellRenderer(new ListCellRenderer<Floor>() {
            private final FloorListRenderer base = new FloorListRenderer();
            @Override
            public Component getListCellRendererComponent(JList<? extends Floor> list,
                                                          Floor value, int index,
                                                          boolean isSelected, boolean cellHasFocus) {
                Component c = base.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (!isSelected && c instanceof JComponent) {
                    boolean highlight = micro_hasAnyOnFloor(value);
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
        floorPanel.add(new JScrollPane(floorList), BorderLayout.CENTER);

        JButton selectForFloorBtn = new JButton("Проставить галочки на этаже (жилые)");
        selectForFloorBtn.addActionListener(e -> selectAllOnCurrentFloorForResidential());
        JPanel btnWrap = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btnWrap.add(selectForFloorBtn);
        floorPanel.add(btnWrap, BorderLayout.SOUTH);

        left.add(secPanel, BorderLayout.NORTH);
        left.add(floorPanel, BorderLayout.CENTER);
        return left;
    }

    private JPanel buildCenterPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new TitledBorder("Помещения"));

        spaceTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        spaceTable.getTableHeader().setFont(HEADER_FONT);
        spaceTable.setRowHeight(25);
        spaceTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) updateRoomList();
        });

        // Подсветка помещений, где есть хотя бы одна галочка микроклимата
        spaceTable.getColumnModel().getColumn(0).setCellRenderer(new javax.swing.table.DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value,
                                                           boolean isSelected, boolean hasFocus,
                                                           int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (!isSelected) {
                    Space s = spaceTableModel.getAt(row);
                    boolean highlight = micro_hasAnyInSpace(s);
                    c.setBackground(highlight ? HL_GREEN : UIManager.getColor("Table.background"));
                }
                return c;
            }
        });

        panel.add(new JScrollPane(spaceTable), BorderLayout.CENTER);
        return panel;
    }


    private JPanel buildRightPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new TitledBorder("Комнаты"));

        roomTable.setRowHeight(28);
        roomTable.getTableHeader().setFont(HEADER_FONT);
        // Устанавливаем кастомный редактор/рендерер для колонки «Наружные стены»
        roomTable.getColumnModel().getColumn(2).setCellRenderer(new WallsButtonsCell());
        roomTable.getColumnModel().getColumn(2).setCellEditor(new WallsButtonsCell());

        panel.add(new JScrollPane(roomTable), BorderLayout.CENTER);
        return panel;
    }

    // ==== Обновления списков ====

    private void refreshSections() {
        sectionListModel.clear();
        if (currentBuilding == null) return;
        currentBuilding.getSections().stream()
                .sorted(SECTION_ORDER)
                .forEach(sectionListModel::addElement);
        if (!sectionListModel.isEmpty()) sectionList.setSelectedIndex(0);
    }

    private void refreshFloors() {
        floorListModel.clear();
        if (currentBuilding == null) return;
        Section selected = sectionList.getSelectedValue();
        if (selected == null) return;
        int secIdx = sectionList.getSelectedIndex();

        // показываем только Жилые, Офисные и Смешанные; скрываем Общественные и Улица
        List<Floor> floors = currentBuilding.getFloors().stream()
                .filter(f -> f.getSectionIndex() == secIdx)
                .filter(f -> {
                    Floor.FloorType t = f.getType();
                    return t == Floor.FloorType.RESIDENTIAL
                            || t == Floor.FloorType.OFFICE
                            || t == Floor.FloorType.MIXED;
                })
                .sorted(FLOOR_ORDER)
                .collect(Collectors.toList());

        floors.forEach(floorListModel::addElement);
        if (!floorListModel.isEmpty()) {
            floorList.setSelectedIndex(0);
        }

        updateSpaceList();
        updateRoomList();
    }


    private void updateSpaceList() {
        spaceTableModel.clear();
        Floor f = floorList.getSelectedValue();
        if (f == null) return;
        List<Space> spaces = new ArrayList<>(f.getSpaces());
        spaces.sort(SPACE_ORDER);
        spaceTableModel.setSpaces(spaces);
        if (spaceTable.getRowCount() > 0) spaceTable.setRowSelectionInterval(0, 0);
    }

    private void updateRoomList() {
        roomTableModel.clear();
        Space s = getSelectedSpace();
        if (s == null) return;

        List<Room> rooms = new ArrayList<>(s.getRooms());
        rooms.sort(ROOM_ORDER);
        roomTableModel.setRooms(rooms);
    }

    private Space getSelectedSpace() {
        int i = spaceTable.getSelectedRow();
        if (i < 0) return null;
        return spaceTableModel.getAt(i);
    }

    // ==== Массовая операция ====

    /** Проставляем ВСЕ чекбоксы на выбранном этаже, только в помещениях APARTMENT. */
    private void selectAllOnCurrentFloorForResidential() {
        Floor f = floorList.getSelectedValue();
        if (f == null) return;

        for (Space s : f.getSpaces()) {
            if (s.getType() == Space.SpaceType.APARTMENT) {
                for (Room r : s.getRooms()) {
                    globalRoomSelectionMap.put(r.getId(), true);
                    r.setSelected(true); // сразу пишем и в модель
                }
            }
        }
        roomTableModel.fireTableDataChanged();
        micro_refreshHighlights();
    }

    // ==== Вспомогательные модели/клетки ====

    /** Таблица помещений (минимум: идентификатор и тип). */
    private static final class SpaceTableModel extends javax.swing.table.AbstractTableModel {
        private final String[] NAMES = {"Помещение", "Тип"};
        private final Class<?>[] TYPES = {String.class, String.class};
        private final List<Space> spaces = new ArrayList<>();

        void setSpaces(List<Space> list) { spaces.clear(); spaces.addAll(list); fireTableDataChanged(); }
        void clear() { spaces.clear(); fireTableDataChanged(); }
        Space getAt(int row) { return spaces.get(row); }

        @Override public int getRowCount() { return spaces.size(); }
        @Override public int getColumnCount() { return NAMES.length; }
        @Override public String getColumnName(int c) { return NAMES[c]; }
        @Override public Class<?> getColumnClass(int c) { return TYPES[c]; }
        @Override public Object getValueAt(int row, int col) {
            Space s = spaces.get(row);
            return (col == 0) ? s.getIdentifier() : String.valueOf(s.getType());
        }
    }
    /** Выбрать секцию по индексу с обновлением этажей/помещений. */
    public void selectSectionByIndex(int idx) {
        if (idx < 0 || idx >= sectionListModel.getSize()) return;
        sectionList.setSelectedIndex(idx);
        // refreshFloors() вызовется через ListSelectionListener;
        // на всякий случай обновим вручную:
        refreshFloors();
    }
    private JPanel buildBottomPanel() {
        JPanel p = new JPanel(new BorderLayout());
        JPanel tempPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 5));
        JLabel tempLabel = new JLabel("Температура воздуха:");
        JRadioButton coldButton = new JRadioButton("Холодно");
        JRadioButton normalButton = new JRadioButton("Нормально");
        JRadioButton warmButton = new JRadioButton("Тепло");
        ButtonGroup tempGroup = new ButtonGroup();
        tempGroup.add(coldButton);
        tempGroup.add(normalButton);
        tempGroup.add(warmButton);
        coldButton.setSelected(true);

        ActionListener tempListener = e -> {
            if (coldButton.isSelected()) {
                temperatureMode = MicroclimateExcelExporter.TemperatureMode.COLD;
            } else if (normalButton.isSelected()) {
                temperatureMode = MicroclimateExcelExporter.TemperatureMode.NORMAL;
            } else if (warmButton.isSelected()) {
                temperatureMode = MicroclimateExcelExporter.TemperatureMode.WARM;
            }
        };
        coldButton.addActionListener(tempListener);
        normalButton.addActionListener(tempListener);
        warmButton.addActionListener(tempListener);

        tempPanel.add(tempLabel);
        tempPanel.add(coldButton);
        tempPanel.add(normalButton);
        tempPanel.add(warmButton);

        JPanel exportPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton exportBtn = new JButton("Экспорт в Excel");

        exportBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        exportBtn.setBackground(new Color(0, 100, 0));
        exportBtn.setForeground(Color.WHITE);
        exportBtn.setFocusPainted(false);
        exportBtn.addActionListener(e -> {
            commitEditors();                 // зафиксировать редактирование таблиц
            updateRoomSelectionStates();     // перелить чекбоксы в модель Room
            MicroclimateExcelExporter.export(
                    currentBuilding,
                    sectionIndexToExport(),  // текущая секция или все
                    temperatureMode,
                    this
            );
        });
        // hover
        exportBtn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) { exportBtn.setBackground(new Color(0, 80, 0)); }
            public void mouseExited (java.awt.event.MouseEvent e) { exportBtn.setBackground(new Color(0,100,0)); }
        });

        exportPanel.add(exportBtn);
        p.add(tempPanel, BorderLayout.WEST);
        p.add(exportPanel, BorderLayout.EAST);
        return p;
    }

    private void commitEditors() {
        // аккуратно завершаем редактирование в обеих таблицах
        JTable[] tables = { spaceTable, roomTable };
        for (JTable t : tables) {
            if (t.isEditing()) {
                try { t.getCellEditor().stopCellEditing(); } catch (Exception ignore) {}
            }
        }
    }

    private int sectionIndexToExport() {
        if (currentBuilding == null || currentBuilding.getSections() == null) return -1;
        if (currentBuilding.getSections().size() <= 1) return -1; // одна секция → экспорт всех (логика экспортера)
        int idx = sectionList.getSelectedIndex();
        return (idx >= 0) ? idx : -1; // если ничего не выбрано — все
    }


    /** Есть ли в помещении хотя бы одна комната с галочкой микроклимата? */
    private boolean micro_hasAnyInSpace(Space s) {
        if (s == null) return false;
        for (Room r : s.getRooms()) {
            Boolean v = globalRoomSelectionMap.get(r.getId());
            boolean chosen = (v != null) ? v : r.isMicroclimateSelected();
            if (chosen) return true;
        }
        return false;
    }

    /** Есть ли на этом этаже хотя бы одно помещение с галочкой микроклимата? */
    private boolean micro_hasAnyOnFloor(Floor f) {
        if (f == null) return false;
        for (Space s : f.getSpaces()) {
            if (micro_hasAnyInSpace(s)) return true;
        }
        return false;
    }

    /** Перерисовать списки/таблицы с учётом подсветки. */
    private void micro_refreshHighlights() {
        try { if (floorList != null) floorList.repaint(); } catch (Throwable ignore) {}
        try { if (spaceTable != null) spaceTable.repaint(); } catch (Throwable ignore) {}
    }

}
