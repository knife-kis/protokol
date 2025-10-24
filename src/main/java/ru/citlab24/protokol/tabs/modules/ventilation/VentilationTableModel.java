package ru.citlab24.protokol.tabs.modules.ventilation;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Табличная модель для вкладки "Вентиляция" с опциональной первой колонкой "Блок-секция".
 */
public class VentilationTableModel extends AbstractTableModel {

    private final List<VentilationRecord> records = new ArrayList<>();
    private boolean showSectionColumn = false;

    /** Включить/выключить первую колонку "Блок-секция". Включай, если секций > 1. */
    public void setShowSectionColumn(boolean show) {
        if (this.showSectionColumn != show) {
            this.showSectionColumn = show;
            fireTableStructureChanged(); // пересоздаёт колонки у JTable
        }
    }
    public boolean isShowSectionColumn() { return showSectionColumn; }

    public void clearData() {
        int sz = records.size();
        if (sz > 0) {
            records.clear();
            fireTableRowsDeleted(0, sz - 1);
        }
    }
    public void addRecord(VentilationRecord r) {
        int idx = records.size();
        records.add(r);
        fireTableRowsInserted(idx, idx);
    }
    public void addRecords(Collection<VentilationRecord> toAdd) {
        if (toAdd == null || toAdd.isEmpty()) return;
        int from = records.size();
        records.addAll(toAdd);
        fireTableRowsInserted(from, records.size() - 1);
    }
    public List<VentilationRecord> getRecords() { return new ArrayList<>(records); }

    // ===== AbstractTableModel =====
    @Override public int getRowCount() { return records.size(); }
    @Override public int getColumnCount() { return showSectionColumn ? 9 : 8; }

    @Override
    public String getColumnName(int column) {
        if (showSectionColumn) {
            return switch (column) {
                case 0 -> "Блок-секция";
                case 1 -> "Этаж";
                case 2 -> "Помещение";
                case 3 -> "Комната";
                case 4 -> "Каналы";
                case 5 -> "Форма";
                case 6 -> "Ширина (м)";
                case 7 -> "Сечение (кв.м)";
                case 8 -> "Объем (куб.м)";
                default -> "";
            };
        } else {
            return switch (column) {
                case 0 -> "Этаж";
                case 1 -> "Помещение";
                case 2 -> "Комната";
                case 3 -> "Каналы";
                case 4 -> "Форма";
                case 5 -> "Ширина (м)";
                case 6 -> "Сечение (кв.м)";
                case 7 -> "Объем (куб.м)";
                default -> "";
            };
        }
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        if (showSectionColumn) {
            return switch (columnIndex) {
                case 0 -> Integer.class;
                case 1,2,3 -> String.class;
                case 4 -> Integer.class;
                case 5 -> VentilationRecord.DuctShape.class;
                case 6,7 -> Double.class;
                case 8 -> Double.class;
                default -> Object.class;
            };
        } else {
            return switch (columnIndex) {
                case 0,1,2 -> String.class;
                case 3 -> Integer.class;
                case 4 -> VentilationRecord.DuctShape.class;
                case 5,6 -> Double.class;
                case 7 -> Double.class;
                default -> Object.class;
            };
        }
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        // редактируем: Каналы, Форма, Ширина, Объём — Сечение запрещено (только вычисляется)
        if (showSectionColumn) return columnIndex == 4 || columnIndex == 5 || columnIndex == 6 || columnIndex == 8;
        else return columnIndex == 3 || columnIndex == 4 || columnIndex == 5 || columnIndex == 7;

    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        VentilationRecord r = records.get(rowIndex);
        if (showSectionColumn) {
            return switch (columnIndex) {
                case 0 -> r.sectionIndex() + 1;
                case 1 -> r.floor();
                case 2 -> r.space();
                case 3 -> r.room();
                case 4 -> r.channels();
                case 5 -> r.shape();
                case 6 -> r.width();
                case 7 -> r.sectionArea();
                case 8 -> r.volume();
                default -> "";
            };
        } else {
            return switch (columnIndex) {
                case 0 -> r.floor();
                case 1 -> r.space();
                case 2 -> r.room();
                case 3 -> r.channels();
                case 4 -> r.shape();
                case 5 -> r.width();
                case 6 -> r.sectionArea();
                case 7 -> r.volume();
                default -> "";
            };
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        VentilationRecord r = records.get(rowIndex);
        try {
            int areaCol  = showSectionColumn ? 7 : 6; // индекс «Сечение»
            if (showSectionColumn) {
                if (columnIndex == 4) { // Каналы
                    int v = toInt(aValue, r.channels());
                    var nr = r.withChannels(v);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationChannels(v);
                    fireTableCellUpdated(rowIndex, columnIndex);
                } else if (columnIndex == 5) { // Форма
                    var shape = (aValue instanceof VentilationRecord.DuctShape ds) ? ds
                            : VentilationRecord.DuctShape.valueOf(aValue.toString());
                    var nr = r.withShape(shape); // площадь пересчиталась
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(nr.sectionArea());
                    fireTableCellUpdated(rowIndex, columnIndex);
                    fireTableCellUpdated(rowIndex, areaCol); // обновить «Сечение»
                } else if (columnIndex == 6) { // Ширина
                    double w = toDouble(aValue, r.width());
                    var nr = r.withWidth(w);    // площадь пересчиталась
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(nr.sectionArea());
                    fireTableCellUpdated(rowIndex, columnIndex);
                    fireTableCellUpdated(rowIndex, areaCol);
                } else if (columnIndex == 8) { // Объём
                    Double v = toNullableDouble(aValue, r.volume());
                    var nr = r.withVolume(v);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVolume(v);
                    fireTableCellUpdated(rowIndex, columnIndex);
                }
            } else {
                if (columnIndex == 3) { // Каналы
                    int v = toInt(aValue, r.channels());
                    var nr = r.withChannels(v);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationChannels(v);
                    fireTableCellUpdated(rowIndex, columnIndex);
                } else if (columnIndex == 4) { // Форма
                    var shape = (aValue instanceof VentilationRecord.DuctShape ds) ? ds
                            : VentilationRecord.DuctShape.valueOf(aValue.toString());
                    var nr = r.withShape(shape);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(nr.sectionArea());
                    fireTableCellUpdated(rowIndex, columnIndex);
                    fireTableCellUpdated(rowIndex, areaCol);
                } else if (columnIndex == 5) { // Ширина
                    double w = toDouble(aValue, r.width());
                    var nr = r.withWidth(w);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(nr.sectionArea());
                    fireTableCellUpdated(rowIndex, columnIndex);
                    fireTableCellUpdated(rowIndex, areaCol);
                } else if (columnIndex == 7) { // Объём
                    Double v = toNullableDouble(aValue, r.volume());
                    var nr = r.withVolume(v);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVolume(v);
                    fireTableCellUpdated(rowIndex, columnIndex);
                }
            }
        } catch (Exception ignore) {}
    }

    // ===== утилиты парсинга =====
    private static int toInt(Object v, int def) {
        if (v == null) return def;
        if (v instanceof Number n) return n.intValue();
        return Integer.parseInt(v.toString().trim());
    }
    private static double toDouble(Object v, double def) {
        if (v == null) return def;
        if (v instanceof Number n) return n.doubleValue();
        return Double.parseDouble(v.toString().trim().replace(",", "."));
    }
    private static Double toNullableDouble(Object v, Double def) {
        if (v == null) return null;
        if (v instanceof Number n) return n.doubleValue();
        String s = v.toString().trim();
        if (s.isEmpty()) return null;
        return Double.parseDouble(s.replace(",", "."));
    }
    public VentilationRecord getRecordAt(int row) { return records.get(row); }
    public void setRecordAt(int row, VentilationRecord r) {
        records.set(row, r);
        if (r.roomRef() != null) r.roomRef().setVentilationSectionArea(r.sectionArea());
        fireTableRowsUpdated(row, row);
    }

}
