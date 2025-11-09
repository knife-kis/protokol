package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Space;
import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.util.*;
import java.util.List;

/**
 * Вкладка «Осв улица».
 * Собирает все комнаты со всех помещений на этажах типа STREET и
 * показывает их в таблице: слева «Название», далее 4 редактируемых значения.
 * Экспорт пока заглушка (кнопка есть, повесим реальный экспорт позже).
 */
public final class StreetLightingTab extends JPanel {
    /** Текущие строки (нужны для снапшота/восстановления значений). */
    private java.util.List<Row> currentRows = new java.util.ArrayList<>();

    private Building building;

    private final RowsTableModel model = new RowsTableModel();
    private final JTable table = new JTable(model);

    public StreetLightingTab(Building building) {
        this.building = (building != null) ? building : new Building();
        setLayout(new BorderLayout(10, 10));
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        add(buildCenterPanel(), BorderLayout.CENTER);
        add(buildBottomBar(), BorderLayout.SOUTH);

        rebuildFromBuilding();
        configureTable();
    }

    public void setBuilding(Building b) {
        this.building = (b != null) ? b : new Building();
        rebuildFromBuilding(); // заново собрать список комнат с этажей типа STREET
    }

    /** Обновить таблицу заново. */
    public void refreshData() {
        rebuildFromBuilding();
    }

    // ============ UI ============

    private JComponent buildCenterPanel() {
        JPanel content = new JPanel(new BorderLayout());
        content.setBorder(titled("Освещение улица — перечень точек"));

        JLabel hint = new JLabel("Собраны все комнаты с этажей типа «Улица». Заполните 4 значения напротив каждого названия.");
        hint.setBorder(BorderFactory.createEmptyBorder(0, 0, 8, 0));

        JPanel top = new JPanel(new BorderLayout());
        top.add(hint, BorderLayout.WEST);

        content.add(top, BorderLayout.NORTH);
        content.add(new JScrollPane(table), BorderLayout.CENTER);
        return content;
    }

    private JComponent buildBottomBar() {
        JPanel bar = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));

        JButton export = new JButton("Экспорт в Excel");
        export.addActionListener(e -> {
            try {
                java.util.List<StreetLightingExcelExporter.RowData> rows = collectRowsForExport();
                StreetLightingExcelExporter.export(rows, this);
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Ошибка экспорта: " + ex.getMessage(),
                        "Ошибка", JOptionPane.ERROR_MESSAGE);
            }
        });

        bar.add(export);
        return bar;
    }


    private TitledBorder titled(String title) {
        return BorderFactory.createTitledBorder(null, title, TitledBorder.LEFT, TitledBorder.TOP);
    }

    private void configureTable() {
        table.setRowHeight(26);
        table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        table.setAutoResizeMode(JTable.AUTO_RESIZE_SUBSEQUENT_COLUMNS);

        // ширины
        if (table.getColumnModel().getColumnCount() >= 5) {
            table.getColumnModel().getColumn(0).setPreferredWidth(380);
            for (int c = 1; c <= 4; c++) {
                table.getColumnModel().getColumn(c).setPreferredWidth(110);
            }
        }

        // единый редактор: один клик для старта + автовыделение текста
        final JTextField tf = new JTextField();
        tf.addFocusListener(new java.awt.event.FocusAdapter() {
            @Override public void focusGained(java.awt.event.FocusEvent e) {
                SwingUtilities.invokeLater(tf::selectAll);
            }
        });
        final DefaultCellEditor editor = new DefaultCellEditor(tf);
        editor.setClickCountToStart(1);

        for (int c = 1; c <= 4; c++) {
            table.getColumnModel().getColumn(c).setCellEditor(editor);
        }

        // фиксировать значение при потере фокуса
        table.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);

        // принудительно запускать редактор по первому клику
        table.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override public void mousePressed(java.awt.event.MouseEvent e) {
                if (e.getClickCount() == 1) {
                    int row = table.rowAtPoint(e.getPoint());
                    int col = table.columnAtPoint(e.getPoint());
                    if (row >= 0 && col >= 1 && table.editCellAt(row, col)) {
                        Component comp = table.getEditorComponent();
                        if (comp != null) {
                            comp.requestFocus();
                            if (comp instanceof JTextField f) f.selectAll();
                        }
                    }
                }
            }
        });
    }

    // ============ ДАННЫЕ ============

    /** Перестроить список строк из текущего Building. */
    private void rebuildFromBuilding() {
        List<Row> rows = new ArrayList<>();
        if (building != null) {
            List<Floor> streetFloors = new ArrayList<>();
            for (Floor f : building.getFloors()) {
                if (f != null && f.getType() == Floor.FloorType.STREET) streetFloors.add(f);
            }
            streetFloors.sort(Comparator.comparingInt(Floor::getPosition));

            for (Floor f : streetFloors) {
                List<Space> spaces = new ArrayList<>(f.getSpaces());
                spaces.sort(Comparator.comparingInt(Space::getPosition));
                for (Space s : spaces) {
                    List<Room> rooms = new ArrayList<>(s.getRooms());
                    rooms.sort(Comparator.comparingInt(Room::getPosition));
                    for (Room r : rooms) {
                        Row row = new Row();
                        String floorPart = safe(f.getNumber());
                        String spacePart = safe(s.getIdentifier());
                        String roomPart  = safe(r.getName());
                        row.displayName  = buildDisplayName(floorPart, spacePart, roomPart); // сейчас это просто roomPart
                        row.key = f.getSectionIndex() + "|" + floorPart + "|" + spacePart + "|" + roomPart;
                        rows.add(row);
                    }
                }
            }
        }
        this.currentRows = rows;     // ← сохраняем ссылку
        model.setRows(rows);
    }

    private static String buildDisplayName(String floor, String space, String room) {
        return room.isEmpty() ? "(без названия)" : room;
    }

    private static String safe(String s) { return (s == null) ? "" : s.trim(); }

    // ============ МОДЕЛЬ ТАБЛИЦЫ ============

    private static final class Row {
        String key;           // устойчивый ключ (секция|этаж|помещение|комната)
        String displayName;   // то, что показываем в первом столбце
        Double v1, v2, v3, v4; // 4 произвольных значения
    }

    private static final class RowsTableModel extends AbstractTableModel {
        private final String[] NAMES = { "Название", "Макс слева", "Мин центр", "Макс справа", "Мин снизу" };
        private final Class<?>[] CLZ = { String.class, Object.class, Object.class, Object.class, Object.class };
        private final java.util.List<Row> rows = new ArrayList<>();

        void setRows(List<Row> newRows) {
            rows.clear();
            if (newRows != null) rows.addAll(newRows);
            fireTableDataChanged();
        }

        @Override public int getRowCount() { return rows.size(); }
        @Override public int getColumnCount() { return NAMES.length; }
        @Override public String getColumnName(int column) { return NAMES[column]; }
        @Override public Class<?> getColumnClass(int columnIndex) { return CLZ[columnIndex]; }
        @Override public boolean isCellEditable(int rowIndex, int columnIndex) { return columnIndex >= 1; }

        @Override public Object getValueAt(int rowIndex, int columnIndex) {
            Row r = rows.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> r.displayName;
                case 1 -> r.v1;
                case 2 -> r.v2;
                case 3 -> r.v3;
                case 4 -> r.v4;
                default -> null;
            };
        }

        @Override public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
            Row r = rows.get(rowIndex);
            Double parsed = parseNumberOrNull(aValue);
            switch (columnIndex) {
                case 1 -> r.v1 = parsed;
                case 2 -> r.v2 = parsed;
                case 3 -> r.v3 = parsed;
                case 4 -> r.v4 = parsed;
            }
            fireTableCellUpdated(rowIndex, columnIndex);
        }

        private static Double parseNumberOrNull(Object v) {
            if (v == null) return null;
            if (v instanceof Number n) return n.doubleValue();
            String s = v.toString().trim();
            if (s.isEmpty()) return null;
            s = s.replace(',', '.');
            try {
                return Double.parseDouble(s);
            } catch (NumberFormatException ex) {
                return null;
            }
        }
    }
    /** Преобразуем текущую таблицу во вход для экспортёра. */
    /** Преобразуем текущую таблицу во вход для экспортёра. */
    private java.util.List<StreetLightingExcelExporter.RowData> collectRowsForExport() {
        java.util.List<StreetLightingExcelExporter.RowData> out = new java.util.ArrayList<>();
        for (int r = 0; r < model.getRowCount(); r++) {
            String name = String.valueOf(model.getValueAt(r, 0));
            Double v1 = coerceToDouble(model.getValueAt(r, 1));
            Double v2 = coerceToDouble(model.getValueAt(r, 2));
            Double v3 = coerceToDouble(model.getValueAt(r, 3));
            Double v4 = coerceToDouble(model.getValueAt(r, 4));
            out.add(new StreetLightingExcelExporter.RowData(name, v1, v2, v3, v4));
        }
        return out;
    }
    /** Снимок значений по ключу "секция|этаж|помещение|комната". */
    public java.util.Map<String, Double[]> snapshotValuesByKey() {
        java.util.Map<String, Double[]> out = new java.util.HashMap<>();
        for (Row r : currentRows) {
            out.put(r.key, new Double[]{ r.v1, r.v2, r.v3, r.v4 });
        }
        return out;
    }

    /** Применить значения по ключу. Если ключ не найден — строка остаётся пустой. */
    public void applyValuesByKey(java.util.Map<String, Double[]> byKey) {
        if (byKey == null || byKey.isEmpty()) return;
        for (Row r : currentRows) {
            Double[] v = byKey.get(r.key);
            if (v != null) {
                r.v1 = (v.length > 0) ? v[0] : null;
                r.v2 = (v.length > 1) ? v[1] : null;
                r.v3 = (v.length > 2) ? v[2] : null;
                r.v4 = (v.length > 3) ? v[3] : null;
            }
        }
        model.fireTableDataChanged();
    }

    /** Безопасно приводим Object к Double (без авто-разупаковки null). */
    private static Double coerceToDouble(Object v) {
        if (v == null) return null;
        if (v instanceof Number n) return Double.valueOf(n.doubleValue()); // всегда объектный Double
        return StreetLightingExcelExporter.tryParse(v);                    // уже возвращает Double или null
    }


}
