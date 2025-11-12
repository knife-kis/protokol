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
    private boolean suppressApplyDialog = false;

    private VentilationRecord lastSnapshot;
    private int lastRow = -1, lastCol = -1;

    private Building building;
    private final VentilationTableModel tableModel = new VentilationTableModel();
    private final JTable ventilationTable = new JTable(tableModel);
    // === Зебра по группам (квартира/офис) ===
    private static final Color STRIPE_A = Color.WHITE;
    private static final Color STRIPE_B = new Color(240, 248, 240); // мягкий зелёно-серый

    /** Порядковый номер «полосы» для группы (section|floor|space) */
    private final java.util.Map<String, Integer> groupStripeIndex = new java.util.LinkedHashMap<>();

    // Сохраняем «родные» рендереры до переустановки,
    private javax.swing.table.TableCellRenderer baseObjRenderer;
    private javax.swing.table.TableCellRenderer baseStrRenderer;
    private javax.swing.table.TableCellRenderer baseNumRenderer;
    private javax.swing.table.TableCellRenderer baseIntRenderer;
    private javax.swing.table.TableCellRenderer baseDblRenderer;
    private javax.swing.table.TableCellRenderer baseEnumRenderer;

    private static final List<String> TARGET_ROOMS = RoomUtils.RESIDENTIAL_ROOM_KEYWORDS;
    private static final List<String> TARGET_FLOORS = List.of("жилой", "смешанный", "офисный", "общественный");

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

        // Стили заголовка
        JTableHeader header = ventilationTable.getTableHeader();
        header.setBackground(new Color(70, 130, 180)); // SteelBlue
        header.setForeground(Color.WHITE);
        header.setFont(new Font("Segoe UI", Font.BOLD, 14));
        header.setReorderingAllowed(false);
        ((DefaultTableCellRenderer) header.getDefaultRenderer()).setHorizontalAlignment(JLabel.CENTER);

        // Не завершать редактор при потере фокуса
        ventilationTable.putClientProperty("terminateEditOnFocusLost", Boolean.FALSE);

        // Снимок до правки (для диалогов массового применения)
        ventilationTable.getSelectionModel().addListSelectionListener(e -> captureSnapshot());
        ventilationTable.getColumnModel().getSelectionModel().addListSelectionListener(e -> captureSnapshot());

        tableModel.addTableModelListener(evt -> {
            if (suppressApplyDialog) return; // массовое применение — молчим
            if (evt.getType() != javax.swing.event.TableModelEvent.UPDATE) return;

            int row = evt.getFirstRow();
            int col = evt.getColumn();
            if (row < 0 || col < 0) return;

            int off = tableModel.isShowSectionColumn() ? 1 : 0;
            int shapeCol = 4 + off;
            int widthCol = 5 + off;

            if (col == shapeCol || col == widthCol) {
                VentilationRecord changed = tableModel.getRecordAt(row);

                Object[] options = {"Да", "Нет", "Отмена"};
                int choice = JOptionPane.showOptionDialog(
                        this,
                        (col == shapeCol
                                ? "Сделать такую ФОРМУ для всех одноименных комнат («" + changed.room() + "»)?"
                                : "Сделать такую ШИРИНУ для всех одноименных комнат («" + changed.room() + "»)?"),
                        "Применить к одноимённым",
                        JOptionPane.YES_NO_CANCEL_OPTION,
                        JOptionPane.QUESTION_MESSAGE,
                        null, options, options[1] // по умолчанию — «Нет»
                );

                if (choice == JOptionPane.CANCEL_OPTION) {
                    if (lastSnapshot != null && lastRow == row) tableModel.setRecordAt(row, lastSnapshot);
                    return;
                }
                if (choice == JOptionPane.YES_OPTION) {
                    suppressApplyDialog = true;
                    try {
                        String key = changed.room() == null ? "" : changed.room().trim();
                        for (int r = 0; r < tableModel.getRowCount(); r++) {
                            if (r == row) continue;
                            VentilationRecord cur = tableModel.getRecordAt(r);
                            if (cur.room() != null && cur.room().trim().equalsIgnoreCase(key)) {
                                if (col == shapeCol) tableModel.setRecordAt(r, cur.withShape(changed.shape()));
                                else                  tableModel.setRecordAt(r, cur.withWidth(changed.width()));
                            }
                        }
                    } finally {
                        suppressApplyDialog = false;
                    }
                }
            }
        });

        // Устанавливаем рендереры «зебры по группам»
        installGroupStripeRenderers();

        add(new JScrollPane(ventilationTable), BorderLayout.CENTER);
        add(createButtonPanel(), BorderLayout.SOUTH);
    }

    // Стало
    private void configureEditorsAndWidths() {
        boolean withSection = tableModel.isShowSectionColumn();

        int channelsCol = withSection ? 4 : 3;
        int shapeCol    = channelsCol + 1; // 5 / 4
        int widthCol    = channelsCol + 2; // 6 / 5
        int areaCol     = channelsCol + 3; // 7 / 6 (НЕ редактируем)
        int volumeCol   = channelsCol + 4; // 8 / 7

        // Каналы (целое) — текстовый ввод
        ventilationTable.getColumnModel().getColumn(channelsCol)
                .setCellEditor(new NumberTextEditor());

        // Форма — ОДИН КЛИК: кастомный редактор, который сразу переключает и коммитит
        ventilationTable.getColumnModel().getColumn(shapeCol)
                .setCellEditor(new OneClickShapeEditor());

        // Ширина — текстовый ввод
        ventilationTable.getColumnModel().getColumn(widthCol)
                .setCellEditor(new NumberTextEditor());

        // Сечение — без редактора

        // Объём — текстовый ввод
        ventilationTable.getColumnModel().getColumn(volumeCol)
                .setCellEditor(new NumberTextEditor());

        if (withSection) {
            DefaultTableCellRenderer center = new DefaultTableCellRenderer();
            center.setHorizontalAlignment(SwingConstants.CENTER);
            ventilationTable.getColumnModel().getColumn(0).setCellRenderer(center);
            ventilationTable.getColumnModel().getColumn(0).setPreferredWidth(90);
        }
    }

    // Стало
    public void saveCalculationsToModel() {
        commitActiveEditor();

        for (VentilationRecord record : tableModel.getRecords()) {
            Room r = record.roomRef();
            if (r == null) continue;

            r.setVentilationChannels(record.channels());
            r.setVentilationSectionArea(record.sectionArea());
            r.setVolume(record.volume());

            // НОВОЕ: сохраним форму и ширину в Room, если есть соответствующие сеттеры
            try {
                // setVentilationDuctShape(String)
                try {
                    r.getClass().getMethod("setVentilationDuctShape", String.class)
                            .invoke(r, record.shape() == null ? null : record.shape().name());
                } catch (Throwable ignore) {}

                // setVentilationWidth(Double) или setVentilationWidth(double)
                Double w = record.width();
                try {
                    r.getClass().getMethod("setVentilationWidth", Double.class).invoke(r, w);
                } catch (NoSuchMethodException ex) {
                    try {
                        r.getClass().getMethod("setVentilationWidth", double.class)
                                .invoke(r, w == null ? 0.0 : w.doubleValue());
                    } catch (Throwable ignore) {}
                } catch (Throwable ignore) {}
            } catch (Throwable ignore) {}
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

    private void loadVentilationData() {
        tableModel.clearData();
        if (building == null) { System.out.println("Здание не загружено!"); return; }

        for (Floor floor : building.getFloors()) {
            String floorTypeText = floor.getType().toString().toLowerCase(Locale.ROOT);
            boolean floorMatches = containsAny(floorTypeText, TARGET_FLOORS); // теперь включает и «общественный»

            System.out.println("Этаж " + floor.getNumber() + " (тип: " + floorTypeText + ") - " +
                    (floorMatches ? "соответствует" : "не соответствует"));
            if (!floorMatches) continue;

            for (Space space : floor.getSpaces()) {
                String spaceTypeText = space.getType().toString().toLowerCase(Locale.ROOT);

                for (Room room : space.getRooms()) {
                    String roomName = room.getName();
                    String norm = RoomUtils.normalizeRoomName(roomName);

                    boolean roomMatches;
                    switch (floor.getType()) {
                        case RESIDENTIAL:
                            // квартиры: кухни/санузлы/ванны/туалет/уборная — уже в RoomUtils
                            roomMatches = RoomUtils.isResidentialRoom(roomName);
                            break;
                        case OFFICE:
                            // офис: берём санпомещения по тем же словам (туалет/уборная и т.п.)
                            roomMatches = RoomUtils.isResidentialRoom(roomName);
                            break;
                        case PUBLIC:
                            // общественный: берём мусорокамеры и т.п. (любой корень «мусор»)
                            roomMatches = norm.contains("мусор");
                            break;
                        case MIXED:
                            // смешанный: если помещение жилое — как жилые; если общественное — «мусор...»
                            switch (space.getType()) {
                                case APARTMENT:
                                    roomMatches = RoomUtils.isResidentialRoom(roomName);
                                    break;
                                case PUBLIC_SPACE:
                                    roomMatches = norm.contains("мусор");
                                    break;
                                default:
                                    roomMatches = false; // офисные на смешанном — не тащим
                            }
                            break;
                        default:
                            roomMatches = false;
                    }

                    if (!roomMatches) continue;

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

        // === ПЕРЕСЧЁТ «зебры по группам» и перерисовка
        recomputeGroupStripes();
        ventilationTable.repaint();
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
    private String normalizeRoomName(String roomName) {
        return roomName.replaceAll("[\\s\\.-]+", " ").trim().toLowerCase(Locale.ROOT);
    }
    // НОВОЕ: редактор «форма воздуховода» в один клик (переключение круг ↔ квадрат)
    private static final class OneClickShapeEditor extends javax.swing.AbstractCellEditor implements javax.swing.table.TableCellEditor {
        private VentilationRecord.DuctShape value;
        private final javax.swing.JLabel view = new javax.swing.JLabel();
        private long lastToggleAt = 0L; // защита от двойного переключения при дабл-клике

        OneClickShapeEditor() {
            view.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        }

        @Override
        public Object getCellEditorValue() {
            return value;
        }

        @Override
        public java.awt.Component getTableCellEditorComponent(javax.swing.JTable table, Object v, boolean isSelected, int row, int column) {
            VentilationRecord.DuctShape current = (v instanceof VentilationRecord.DuctShape ds)
                    ? ds
                    : VentilationRecord.DuctShape.valueOf(String.valueOf(v));

            long now = System.currentTimeMillis();
            if (now - lastToggleAt < 200) {
                // если второй клик прилетел сразу — не переключаем ещё раз
                value = current;
            } else {
                value = (current == VentilationRecord.DuctShape.CIRCLE)
                        ? VentilationRecord.DuctShape.SQUARE
                        : VentilationRecord.DuctShape.CIRCLE;
                lastToggleAt = now;
            }

            view.setText(value.toString());

            // моментально зафиксировать изменение, чтобы сработал TableModelListener (вопрос «Применить ко всем?»)
            javax.swing.SwingUtilities.invokeLater(this::stopCellEditing);
            return view;
        }
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
    private void captureSnapshot() {
        int r = ventilationTable.getSelectedRow();
        if (r < 0) return;
        lastRow = r;
        lastCol = ventilationTable.getSelectedColumn();
        VentilationRecord x = tableModel.getRecordAt(r);
        if (x != null) {
            lastSnapshot = new VentilationRecord(
                    x.floor(), x.space(), x.room(), x.channels(),
                    x.sectionArea(), x.volume(), x.roomRef(), x.sectionIndex(),
                    x.shape(), x.width()
            );
        }
    }
    // Текстовый редактор чисел: без стрелок, при старте поле очищается.
// Если оставить пусто и подтвердить — значение не меняется.
    private static final class NumberTextEditor extends DefaultCellEditor {
        private Object original;

        NumberTextEditor() {
            super(new JTextField());
            JTextField tf = (JTextField) getComponent();
            tf.setHorizontalAlignment(JTextField.CENTER);
        }
        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
            original = value;
            JTextField tf = (JTextField) getComponent();
            tf.setText("");             // очищаем для быстрого ввода
            tf.requestFocusInWindow();
            return tf;
        }
        @Override
        public Object getCellEditorValue() {
            String s = ((JTextField) getComponent()).getText().trim();
            if (s.isEmpty()) return original;              // пусто → не менять
            return s.replace(',', '.');                    // поддержка запятой
        }
    }
    public Building getBuilding() { return building; }
    // ВНИМАНИЕ: новый метод для общего экспорта (кнопки в "Характеристиках здания")
    public java.util.List<VentilationRecord> getRecordsForExport() {
        commitActiveEditor();        // завершить ввод в ячейке, если редактируется
        saveCalculationsToModel();   // записать channels/sectionArea/volume обратно в Room
        // вернуть копию текущих записей (включая «нулевые» — отфильтруем в экспортере)
        return new java.util.ArrayList<>(tableModel.getRecords());
    }
    public void refreshData() {
        boolean showSection = building != null
                && building.getSections() != null
                && building.getSections().size() > 1;

        tableModel.setShowSectionColumn(showSection);
        configureEditorsAndWidths();

        // после смены структуры безопасно переустановим рендереры
        installGroupStripeRenderers();

        loadVentilationData();       // внутри — пересчёт зебры
    }
    /** Установить рендереры, которые подкрашивают ВСЕ столбцы «полосами» по группам. */
    private void installGroupStripeRenderers() {
        // Сохраняем «родные» делегаты один раз
        baseObjRenderer  = ventilationTable.getDefaultRenderer(Object.class);
        baseStrRenderer  = ventilationTable.getDefaultRenderer(String.class);
        baseNumRenderer  = ventilationTable.getDefaultRenderer(Number.class);
        baseIntRenderer  = ventilationTable.getDefaultRenderer(Integer.class);
        baseDblRenderer  = ventilationTable.getDefaultRenderer(Double.class);
        baseEnumRenderer = ventilationTable.getDefaultRenderer(Enum.class);

        // На все базовые типы ставим наш обёрточный рендерер
        ventilationTable.setDefaultRenderer(Object.class,  new GroupStripeRenderer(baseObjRenderer));
        ventilationTable.setDefaultRenderer(String.class,  new GroupStripeRenderer(baseStrRenderer));
        ventilationTable.setDefaultRenderer(Number.class,  new GroupStripeRenderer(baseNumRenderer));
        ventilationTable.setDefaultRenderer(Integer.class, new GroupStripeRenderer(baseIntRenderer));
        ventilationTable.setDefaultRenderer(Double.class,  new GroupStripeRenderer(baseDblRenderer));
        ventilationTable.setDefaultRenderer(Enum.class,    new GroupStripeRenderer(baseEnumRenderer));
    }

    /** Пересчитать «зебру» так, чтобы все строки одной квартиры/офиса были одного цвета. */
    private void recomputeGroupStripes() {
        groupStripeIndex.clear();
        int band = 0;

        // идём по текущему порядку строк модели: новая группа → следующий цвет
        for (int i = 0; i < tableModel.getRowCount(); i++) {
            VentilationRecord r = tableModel.getRecordAt(i);
            String key = groupKeyForRecord(r);

            if (!groupStripeIndex.containsKey(key)) {
                groupStripeIndex.put(key, band++);
            }
        }
    }

    /** Ключ группы: секция|этаж|помещение (офисы тоже входят) */
    private String groupKeyForRecord(VentilationRecord r) {
        String floor = r.floor() == null ? "" : r.floor();
        String space = r.space() == null ? "" : r.space();
        return r.sectionIndex() + "|" + floor + "|" + space;
    }

    /** Ключ группы по номеру строки модели. */
    private String groupKeyForRow(int modelRow) {
        VentilationRecord r = tableModel.getRecordAt(modelRow);
        return groupKeyForRecord(r);
    }

    /** Обёртка-рендерер: делегирует базовому, но красит фон «полосами по группам» и центрирует текст. */
    private final class GroupStripeRenderer implements javax.swing.table.TableCellRenderer {
        private final javax.swing.table.TableCellRenderer delegate;

        GroupStripeRenderer(javax.swing.table.TableCellRenderer delegate) {
            this.delegate = (delegate != null) ? delegate : new DefaultTableCellRenderer();
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean isSelected, boolean hasFocus,
                                                       int row, int column) {
            // делегируем базовому
            Component c = delegate.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);

            // центрирование для наглядности
            if (c instanceof JLabel lbl) {
                lbl.setHorizontalAlignment(JLabel.CENTER);
            }

            // Определяем цвет полосы по группе
            if (!isSelected) {
                int modelRow = table.convertRowIndexToModel(row);
                String key = groupKeyForRow(modelRow);
                Integer band = groupStripeIndex.get(key);
                Color bg = (band != null && (band % 2 == 1)) ? STRIPE_B : STRIPE_A;
                c.setBackground(bg);
            }

            // Спец-форматирование столбца «Объем (куб.м)» — показывать 1 знак, 0.0 → пусто
            try {
                int volumeCol = tableModel.isShowSectionColumn() ? 8 : 7;
                if (column == volumeCol && c instanceof JLabel lbl) {
                    if (value == null) {
                        lbl.setText("");
                    } else if (value instanceof Number n) {
                        double dv = n.doubleValue();
                        lbl.setText(dv == 0.0 ? "" : String.format(Locale.ROOT, "%.1f", dv));
                    } else {
                        // текст оставляем как есть
                    }
                }
            } catch (Throwable ignore) {}

            return c;
        }
    }

}
