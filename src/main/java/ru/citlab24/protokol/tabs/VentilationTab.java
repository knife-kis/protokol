package ru.citlab24.protokol.tabs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.JTableHeader;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.util.List;
import java.util.Locale;


public class VentilationTab extends JPanel {
    private Building building;
    private final VentilationTableModel tableModel = new VentilationTableModel(this);
    private final JTable ventilationTable = new JTable(tableModel);

    private static final List<String> TARGET_ROOMS = List.of(
            "кухня", "кухня-ниша",
            "санузел", "сан узел", "сан. узел",
            "ванная", "ванная комната",
            "совмещенный", "совмещенный санузел", "туалет"
    );

    private static final List<String> TARGET_FLOORS = List.of(
            "жилой", "смешанный", "офисный", "офисный"
    );

    public VentilationTab(Building building) {
        this.building = building;
        initUI();
        loadVentilationData();
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    private void initUI() {
        setLayout(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        // Настройка таблицы
        ventilationTable.setRowHeight(30);
        ventilationTable.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        ventilationTable.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));

        // Редакторы ячеек
        ventilationTable.getColumnModel().getColumn(3).setCellEditor(
                new SpinnerEditor(1, 1, 300, 1));
        ventilationTable.getColumnModel().getColumn(4).setCellEditor(
                new SpinnerEditor(0.008, 0.001, 0.1, 0.001));
        ventilationTable.getColumnModel().getColumn(5).setCellEditor(
                new SpinnerEditor(0.0, 0.0, 1000.0, 0.1));

        // Стилизация
        ventilationTable.setShowGrid(true);
        ventilationTable.setGridColor(new Color(220, 220, 220));
        ventilationTable.setIntercellSpacing(new Dimension(1, 1));

        // Центрирование содержимого для всех столбцов
        DefaultTableCellRenderer centerRenderer = new DefaultTableCellRenderer();
        centerRenderer.setHorizontalAlignment(JLabel.CENTER);
        for (int i = 0; i < ventilationTable.getColumnCount(); i++) {
            ventilationTable.getColumnModel().getColumn(i).setCellRenderer(centerRenderer);
        }

        // Чередование цветов строк
        ventilationTable.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
            @Override
            public Component getTableCellRendererComponent(JTable table, Object value,
                                                           boolean isSelected, boolean hasFocus,
                                                           int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

                // Обработка пустых значений
                if (value == null) {
                    ((JLabel) c).setText("");
                }

                if (!isSelected) {
                    c.setBackground(row % 2 == 0 ? Color.WHITE : new Color(240, 240, 240));
                }

                ((JLabel) c).setHorizontalAlignment(JLabel.CENTER);
                return c;
            }
        });

        // Стилизация заголовка
        JTableHeader header = ventilationTable.getTableHeader();
        header.setBackground(new Color(70, 130, 180)); // SteelBlue
        header.setForeground(Color.WHITE);
        header.setFont(new Font("Segoe UI", Font.BOLD, 14));
        header.setReorderingAllowed(false);

        DefaultTableCellRenderer headerRenderer = (DefaultTableCellRenderer) header.getDefaultRenderer();
        headerRenderer.setHorizontalAlignment(JLabel.CENTER);

        add(new JScrollPane(ventilationTable), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    public void saveCalculationsToModel() {
        // Сохранение данных обратно в модель
        for (VentilationRecord record : tableModel.getRecords()) {
            record.roomRef().setVentilationChannels(record.channels());
            record.roomRef().setVentilationSectionArea(record.sectionArea());
        }
    }
    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton saveBtn = new JButton("Сохранить расчет");
        JButton exportBtn = new JButton("Экспорт в Excel");
        exportBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        exportBtn.setBackground(new Color(0, 100, 0)); // Темно-зеленый
        exportBtn.setForeground(Color.WHITE);
        exportBtn.setFocusPainted(false);
        exportBtn.addActionListener(e -> exportToExcel());
        saveBtn.setFont(new Font("Segoe UI", Font.BOLD, 14));
        saveBtn.setBackground(new Color(46, 125, 50)); // Зеленый
        saveBtn.setForeground(Color.WHITE);
        saveBtn.setFocusPainted(false);
        saveBtn.addActionListener(e -> saveCalculations());

        // Эффект при наведении
        saveBtn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) {
                saveBtn.setBackground(new Color(35, 110, 40));
            }
            public void mouseExited(java.awt.event.MouseEvent e) {
                saveBtn.setBackground(new Color(46, 125, 50));
            }
        });
        exportBtn.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseEntered(java.awt.event.MouseEvent e) {
                exportBtn.setBackground(new Color(0, 80, 0));
            }
            public void mouseExited(java.awt.event.MouseEvent e) {
                exportBtn.setBackground(new Color(0, 100, 0));
            }
        });

        panel.add(saveBtn);
        panel.add(exportBtn);

        return panel;
    }

    // МЕТОД ЭКСПОРТА ДАННЫХ В EXCEL
    private void exportToExcel() {
        VentilationExcelExporter.export(tableModel.getRecords(), this);
    }
    public void refreshData() {
        // Добавить логирование для отладки
        System.out.println("Обновление данных вентиляции...");
        System.out.println("Ссылка на здание: " + (building != null ? "не null" : "null"));
        if (building != null) {
            System.out.println("Количество этажей: " + building.getFloors().size());
        }

        loadVentilationData();
    }

    private void loadVentilationData() {
        tableModel.clearData();

        if (building == null) {
            System.out.println("Здание не загружено!");
            return;
        }

        // Добавляем типы жилых помещений
        final List<String> RESIDENTIAL_SPACE_TYPES = List.of(
                "квартира", "жилое помещение", "жилая ячейка"
        );

        System.out.println("Загрузка данных вентиляции для здания: " + building.getName());
        System.out.println("Количество этажей: " + building.getFloors().size());

        for (Floor floor : building.getFloors()) {
            String floorType = floor.getType().toString().toLowerCase(Locale.ROOT);
            boolean floorMatches = containsAny(floorType, TARGET_FLOORS);

            System.out.println("Этаж " + floor.getNumber() +
                    " (тип: " + floorType + ") - " +
                    (floorMatches ? "соответствует" : "не соответствует"));

            if (!floorMatches) continue;

            System.out.println("  Помещений на этаже: " + floor.getSpaces().size());
            for (Space space : floor.getSpaces()) {
                String spaceType = space.getType().toString().toLowerCase(Locale.ROOT);
                System.out.println("  Помещение: " + space.getIdentifier() +
                        " (тип: " + spaceType + ")");
                System.out.println("  Комнат в помещении: " + space.getRooms().size());

                // Определяем режим фильтрации для комнат
                boolean filterRooms;
                if (floorType.contains("жилой")) {
                    filterRooms = true;
                } else if (floorType.contains("смешанный")) {
                    // Для смешанного этажа фильтруем только квартиры
                    filterRooms = containsAny(spaceType, RESIDENTIAL_SPACE_TYPES);
                } else {
                    // Для офисного этажа - без фильтрации
                    filterRooms = false;
                }

                for (Room room : space.getRooms()) {
                    String roomName = room.getName();
                    boolean roomMatches;

                    if (filterRooms) {
                        roomMatches = matchesRoomType(roomName);
                    } else {
                        // Без фильтра - добавляем все комнаты
                        roomMatches = true;
                    }

                    System.out.println("    Комната: " + roomName + " - " +
                            (roomMatches ? "соответствует" : "не соответствует"));

                    if (!roomMatches) continue;

                    System.out.println("      >>> ДОБАВЛЕНА В ТАБЛИЦУ");
                    tableModel.addRecord(new VentilationRecord(
                            floor.getNumber(),
                            space.getIdentifier(),
                            room.getName(),
                            room.getVentilationChannels(),
                            room.getVentilationSectionArea(),
                            room.getVolume(),
                            room
                    ));
                }
            }
        }
        System.out.println("Загружено записей: " + tableModel.getRowCount());
        tableModel.fireTableDataChanged();
    }
    private boolean containsAny(String source, List<String> targets) {
        String lowerSource = source.toLowerCase(Locale.ROOT);
        for (String target : targets) {
            if (lowerSource.contains(target)) {
                return true;
            }
        }
        return false;
    }

    private boolean matchesRoomType(String roomName) {
        if (roomName == null) return false;

        String normalized = roomName
                .replaceAll("[\\s.-]+", " ") // Объединение всех замен
                .trim()
                .toLowerCase(Locale.ROOT);

        return TARGET_ROOMS.stream().anyMatch(normalized::contains);
    }

    private void saveCalculations() {
        saveCalculationsToModel();
        JOptionPane.showMessageDialog(this, "Расчеты сохранены успешно!", "Сохранение", JOptionPane.INFORMATION_MESSAGE);
    }

    public String getRoomCategory(String roomName) {
        if (roomName == null) return null;
        String normalized = normalizeRoomName(roomName);
        if (normalized.contains("кухня")) {
            return "кухня";
        } else if (normalized.contains("санузел") ||
                normalized.contains("сан узел") ||
                normalized.contains("туалет") ||
                normalized.contains("совмещенный")) {
            return "санузел";
        } else if (normalized.contains("ванная")) {
            return "ванная";
        }
        return null;
    }

    private String normalizeRoomName(String roomName) {
        return roomName
                .replaceAll("[\\s\\.-]+", " ")
                .trim()
                .toLowerCase(Locale.ROOT);
    }
    public void setBuilding(Building building) {
        this.building = building;
    }

    public Building getBuilding() {
        return building;
    }
}