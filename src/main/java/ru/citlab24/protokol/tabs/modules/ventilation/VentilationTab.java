package ru.citlab24.protokol.tabs.modules.ventilation;

import ru.citlab24.protokol.tabs.SpinnerEditor;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;
import ru.citlab24.protokol.tabs.utils.RoomUtils;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.JTableHeader;
import java.awt.*;
import java.util.List;
import java.util.Locale;

public class VentilationTab extends JPanel {
    private Building building;
    private final VentilationTableModel tableModel = new VentilationTableModel();
    private final JTable ventilationTable = new JTable(tableModel);

    private static final List<String> TARGET_ROOMS = RoomUtils.RESIDENTIAL_ROOM_KEYWORDS;
    private static final List<String> TARGET_FLOORS = List.of("жилой", "смешанный", "офисный");

    public VentilationTab(Building building) {
        initUI();
        setBuilding(building);      // ← важно: сразу включим/выключим колонку "Блок-секция" и редакторы
        loadVentilationData();
    }

    private void initUI() {
        setLayout(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        ventilationTable.setRowHeight(30);
        ventilationTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        ventilationTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));

        // Грид и интервалы
        ventilationTable.setShowGrid(true);
        ventilationTable.setGridColor(new Color(220, 220, 220));
        ventilationTable.setIntercellSpacing(new Dimension(1, 1));

        // Универсальный рендерер: центрирование + зебра + форматирование столбца "Объем"
        ventilationTable.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value,
                                                           boolean isSelected, boolean hasFocus,
                                                           int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

                // чередование строк
                if (!isSelected) {
                    c.setBackground(row % 2 == 0 ? Color.WHITE : new Color(240, 240, 240));
                }
                ((JLabel) c).setHorizontalAlignment(JLabel.CENTER);

                // индекс столбца "Объем" зависит от наличия колонки "Блок-секция"
                int volumeCol = tableModel.isShowSectionColumn() ? 6 : 5;
                if (column == volumeCol) {
                    if (value == null) {
                        ((JLabel) c).setText("");
                    } else if (value instanceof Number d) {
                        double dv = d.doubleValue();
                        ((JLabel) c).setText(dv == 0.0 ? "" : String.format(Locale.ROOT, "%.1f", dv));
                    }
                }
                return c;
            }
        });

        // Стили заголовка
        JTableHeader header = ventilationTable.getTableHeader();
        header.setBackground(new Color(70, 130, 180)); // SteelBlue
        header.setForeground(Color.WHITE);
        header.setFont(new Font("Segoe UI", Font.BOLD, 14));
        header.setReorderingAllowed(false);
        ((DefaultTableCellRenderer) header.getDefaultRenderer()).setHorizontalAlignment(JLabel.CENTER);
        ventilationTable.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);

        add(new JScrollPane(ventilationTable), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    private void configureEditorsAndWidths() {
        // после любого изменения структуры колонок (fireTableStructureChanged)
        int offset = tableModel.isShowSectionColumn() ? 1 : 0;

        ventilationTable.getColumnModel().getColumn(3 + offset).setCellEditor(
                new SpinnerEditor(1, 0, 300, 1));
        ventilationTable.getColumnModel().getColumn(4 + offset).setCellEditor(
                new SpinnerEditor(0.008, 0.001, 0.1, 0.001));
        ventilationTable.getColumnModel().getColumn(5 + offset).setCellEditor(
                new SpinnerEditor(0.0, 0.0, 1000.0, 0.1));

        if (tableModel.isShowSectionColumn()) {
            DefaultTableCellRenderer center = new DefaultTableCellRenderer();
            center.setHorizontalAlignment(SwingConstants.CENTER);
            ventilationTable.getColumnModel().getColumn(0).setCellRenderer(center);
            ventilationTable.getColumnModel().getColumn(0).setPreferredWidth(90);
        }
    }

    public void saveCalculationsToModel() {
        commitActiveEditor(); // ← добавили

        for (VentilationRecord record : tableModel.getRecords()) {
            record.roomRef().setVentilationChannels(record.channels());
            record.roomRef().setVentilationSectionArea(record.sectionArea());
            record.roomRef().setVolume(record.volume());
        }
    }

    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton saveBtn = new JButton("Сохранить расчет");
        JButton exportBtn = new JButton("Экспорт в Excel");

        exportBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        exportBtn.setBackground(new Color(0, 100, 0));
        exportBtn.setForeground(Color.WHITE);
        exportBtn.setFocusPainted(false);
        exportBtn.addActionListener(e -> {
            commitActiveEditor();
            exportToExcel();
        });

        saveBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        saveBtn.setBackground(new Color(46, 125, 50));
        saveBtn.setForeground(Color.WHITE);
        saveBtn.setFocusPainted(false);
        saveBtn.addActionListener(e -> {
            commitActiveEditor();
            saveCalculations();
        });

        // hover
        saveBtn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) { saveBtn.setBackground(new Color(35, 110, 40)); }
            public void mouseExited (java.awt.event.MouseEvent e) { saveBtn.setBackground(new Color(46, 125, 50)); }
        });
        exportBtn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) { exportBtn.setBackground(new Color(0, 80, 0)); }
            public void mouseExited (java.awt.event.MouseEvent e) { exportBtn.setBackground(new Color(0,100,0)); }
        });

        panel.add(saveBtn);
        panel.add(exportBtn);
        return panel;
    }

    // Экспорт
    private void exportToExcel() {
        commitActiveEditor(); // ← добавили
        // опционально: синхронизировать в Room перед экспортом
        saveCalculationsToModel();
        VentilationExcelExporter.export(tableModel.getRecords(), this);
    }


    public void refreshData() {
        System.out.println("Обновление данных вентиляции...");
        System.out.println("Ссылка на здание: " + (building != null ? "не null" : "null"));
        if (building != null) System.out.println("Количество этажей: " + building.getFloors().size());
        loadVentilationData();
    }

    private void loadVentilationData() {
        tableModel.clearData();
        if (building == null) { System.out.println("Здание не загружено!"); return; }

        final List<String> RESIDENTIAL_SPACE_TYPES = List.of("квартира", "жилое помещение", "жилая ячейка");

        System.out.println("Загрузка данных вентиляции для здания: " + building.getName());
        System.out.println("Количество этажей: " + building.getFloors().size());

        for (Floor floor : building.getFloors()) {
            String floorType = floor.getType().toString().toLowerCase(Locale.ROOT);
            boolean floorMatches = containsAny(floorType, TARGET_FLOORS);

            System.out.println("Этаж " + floor.getNumber() + " (тип: " + floorType + ") - " +
                    (floorMatches ? "соответствует" : "не соответствует"));
            if (!floorMatches) continue;

            System.out.println("  Помещений на этаже: " + floor.getSpaces().size());
            for (Space space : floor.getSpaces()) {
                String spaceType = space.getType().toString().toLowerCase(Locale.ROOT);
                System.out.println("  Помещение: " + space.getIdentifier() + " (тип: " + spaceType + ")");
                System.out.println("  Комнат в помещении: " + space.getRooms().size());

                // режим фильтрации
                boolean filterRooms;
                if (floorType.contains("жилой")) {
                    filterRooms = true;
                } else if (floorType.contains("смешанный")) {
                    filterRooms = containsAny(spaceType, RESIDENTIAL_SPACE_TYPES);
                } else {
                    filterRooms = false;
                }

                for (Room room : space.getRooms()) {
                    String roomName = room.getName();
                    boolean roomMatches = filterRooms ? matchesRoomType(roomName) : true;

                    System.out.println("    Комната: " + roomName + " - " + (roomMatches ? "соответствует" : "не соответствует"));
                    if (!roomMatches) continue;

                    System.out.println("      >>> ДОБАВЛЕНА В ТАБЛИЦУ");
                    tableModel.addRecord(new VentilationRecord(
                            floor.getNumber(),
                            space.getIdentifier(),
                            room.getName(),
                            room.getVentilationChannels(),
                            room.getVentilationSectionArea(),
                            (room.getVolume() != null && room.getVolume() == 0.0) ? null : room.getVolume(),
                            room,
                            floor.getSectionIndex()
                    ));
                }
            }
        }
        System.out.println("Загружено записей: " + tableModel.getRowCount());
        tableModel.fireTableDataChanged();
    }

    private boolean containsAny(String source, List<String> targets) {
        String lower = source.toLowerCase(Locale.ROOT);
        for (String t : targets) if (lower.contains(t)) return true;
        return false;
    }
    private boolean matchesRoomType(String roomName) {
        if (roomName == null) return false;
        String normalized = roomName.replaceAll("[\\s.-]+", " ").trim().toLowerCase(Locale.ROOT);
        return TARGET_ROOMS.stream().anyMatch(normalized::contains);
    }

    private void saveCalculations() {
        commitActiveEditor(); // ← добавили
        saveCalculationsToModel();
        JOptionPane.showMessageDialog(this, "Расчеты сохранены успешно! " +
                        "Записи с нулевыми каналами не будут экспортированы в Excel.",
                "Сохранение", JOptionPane.INFORMATION_MESSAGE);
    }


    public String getRoomCategory(String roomName) {
        if (roomName == null) return null;
        String normalized = normalizeRoomName(roomName);
        if (normalized.contains("кухня")) return "кухня";
        if (normalized.contains("санузел") || normalized.contains("сан узел")
                || normalized.contains("туалет") || normalized.contains("совмещенный")) return "санузел";
        if (normalized.contains("ванная")) return "ванная";
        return null;
    }
    private String normalizeRoomName(String roomName) {
        return roomName.replaceAll("[\\s\\.-]+", " ").trim().toLowerCase(Locale.ROOT);
    }

    public void setBuilding(Building building) {
        this.building = building;

        boolean showSection = building != null
                && building.getSections() != null
                && building.getSections().size() > 1;

        tableModel.setShowSectionColumn(showSection);
        configureEditorsAndWidths();  // ← заново настраиваем редакторы/ширины после смены структуры
    }
    private void commitActiveEditor() {
        if (ventilationTable.isEditing()) {
            try {
                ventilationTable.getCellEditor().stopCellEditing();
            } catch (Exception ignore) {
                // на всякий случай
            }
        }
    }


    public Building getBuilding() { return building; }
}
