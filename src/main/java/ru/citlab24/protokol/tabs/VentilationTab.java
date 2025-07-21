package ru.citlab24.protokol.tabs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.JTableHeader;
import java.awt.*;
import java.util.ArrayList;
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
            record.roomRef().setVolume(record.volume());
        }
    }
    private JPanel createButtonPanel() {
        JButton saveBtn = new JButton("Сохранить расчет");
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

        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        panel.add(saveBtn);
        return panel;
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
                System.out.println("  Помещение: " + space.getIdentifier() +
                        " (тип: " + space.getType() + ")");
                System.out.println("  Комнат в помещении: " + space.getRooms().size());

                for (Room room : space.getRooms()) {
                    String roomName = room.getName();
                    boolean roomMatches = matchesRoomType(roomName);

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
        return switch (normalized) {
            case "кухня" -> "кухня";
            case "кухня ниша" -> "кухня-ниша";
            case "санузел", "сан узел", "сан. узел" -> "санузел";
            case "туалет" -> "туалет";
            case "ванная" -> "ванная";
            case "ванная комната" -> "ванная комната";
            case "совмещенный", "совмещенный санузел" -> "совмещенный санузел";
            default -> null;
        };
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

    // Модель таблицы
    private static class VentilationTableModel extends AbstractTableModel {

        private final String[] COLUMN_NAMES = {
                "Этаж", "Помещение", "Комната",
                "Кол-во каналов", "Сечение (кв.м)", "Объем (куб.м)"
        };

        private final VentilationTab ventilationTab;
        private final List<VentilationRecord> records = new ArrayList<>();

        public VentilationTableModel(VentilationTab ventilationTab) {
            this.ventilationTab = ventilationTab;
        }

        public void clearData() {
            records.clear();
        }

        public void addRecord(VentilationRecord record) {
            records.add(record);
        }

        public List<VentilationRecord> getRecords() {
            return records;
        }

        @Override
        public int getRowCount() {
            return records.size();
        }

        @Override
        public int getColumnCount() {
            return COLUMN_NAMES.length;
        }

        @Override
        public String getColumnName(int column) {
            return COLUMN_NAMES[column];
        }

        @Override
        public Class<?> getColumnClass(int columnIndex) {
            return switch (columnIndex) {
                case 0, 1, 2 -> String.class;
                case 3 -> Integer.class;
                case 4 -> Double.class;
                case 5 -> Object.class; // Изменено на Object.class
                default -> Object.class;
            };
        }

        @Override
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            return columnIndex == 3 || columnIndex == 4 || columnIndex == 5;
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            VentilationRecord record = records.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> record.floor();
                case 1 -> record.space();
                case 2 -> record.room();
                case 3 -> record.channels();
                case 4 -> record.sectionArea();
                case 5 -> record.volume(); // Убрано преобразование 0.0 в null
                default -> null;
            };
        }

        @Override
        public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
            if (aValue == null) return;

            VentilationRecord record = records.get(rowIndex);
            switch (columnIndex) {
                case 3: // Кол-во каналов
                    if (aValue instanceof Number) {
                        int newValue = ((Number) aValue).intValue();
                        int oldValue = record.channels();

                        if (newValue == oldValue) return;

                        String category = ventilationTab.getRoomCategory(record.room());

                        // Если категория не определена, применяем только к текущей комнате
                        if (category == null) {
                            updateRecordAndRoom(record, rowIndex, newValue);
                            return;
                        }

                        // Диалог подтверждения
                        int option = JOptionPane.showOptionDialog(
                                ventilationTab,
                                "Вы хотите изменить количество каналов во всех комнатах типа '" + category + "'?",
                                "Подтверждение",
                                JOptionPane.YES_NO_CANCEL_OPTION,
                                JOptionPane.QUESTION_MESSAGE,
                                null,
                                new String[]{"Да", "Нет", "Отмена"},
                                "Нет"
                        );

                        // Обработка выбора
                        switch (option) {
                            case 0: // Да
                                updateAllRoomsOfType(category, newValue, rowIndex);
                                break;
                            case 1: // Нет
                                updateRecordAndRoom(record, rowIndex, newValue);
                                break;
                            case 2: // Отмена
                                records.set(rowIndex, record.withChannels(oldValue));
                                fireTableCellUpdated(rowIndex, columnIndex); // Обновляем без изменений
                                break;
                        }
                    }
                    break;
                case 4: // Сечение
                    if (aValue instanceof Number) {
                        double doubleValue = ((Number) aValue).doubleValue();
                        records.set(rowIndex, record.withSectionArea(doubleValue));
                    }
                    break;
                case 5: // Объем
                    if (aValue instanceof Number) {
                        double doubleValue = ((Number) aValue).doubleValue();
                        records.set(rowIndex, record.withVolume(doubleValue));
                    }
                    break;
                default:
                    return;
            }
            fireTableCellUpdated(rowIndex, columnIndex);
        }

        private void updateAllRoomsOfType(String category, int newValue, int currentRow) {
            // Обновляем все комнаты в здании
            for (Floor floor : ventilationTab.building.getFloors()) {
                for (Space space : floor.getSpaces()) {
                    for (Room room : space.getRooms()) {
                        if (category.equals(ventilationTab.getRoomCategory(room.getName()))) {
                            room.setVentilationChannels(newValue);
                        }
                    }
                }
            }

            // Обновляем записи в таблице (ВСЕ, включая текущую)
            for (int i = 0; i < records.size(); i++) {
                VentilationRecord r = records.get(i);
                if (category.equals(ventilationTab.getRoomCategory(r.room()))) {
                    records.set(i, r.withChannels(newValue));
                    fireTableCellUpdated(i, 3);
                }
            }
        }

        private void updateRecordAndRoom(VentilationRecord record, int rowIndex, int newValue) {
            records.set(rowIndex, record.withChannels(newValue));
            record.roomRef().setVentilationChannels(newValue);
            fireTableCellUpdated(rowIndex, 3);
        }
    }

    // Record с ссылкой на модель
    private record VentilationRecord(
            String floor,
            String space,
            String room,
            int channels,
            double sectionArea,
            Double volume,
            Room roomRef
    ) {
        public VentilationRecord withVolume(Double newVolume) {
            return new VentilationRecord(floor, space, room, channels, sectionArea, newVolume, roomRef);
        }
        public VentilationRecord withChannels(int newChannels) {
            return new VentilationRecord(floor, space, room, newChannels, sectionArea, volume, roomRef);
        }

        public VentilationRecord withSectionArea(double newArea) {
            return new VentilationRecord(floor, space, room, channels, newArea, volume, roomRef);
        }
    }
}