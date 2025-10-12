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
    public VentilationRecord getRecordAt(int row) { return records.get(row); }
    public List<VentilationRecord> getRecords() { return new ArrayList<>(records); }

    // ===== AbstractTableModel =====
    @Override public int getRowCount() { return records.size(); }
    @Override public int getColumnCount() { return showSectionColumn ? 7 : 6; }

    @Override
    public String getColumnName(int column) {
        if (showSectionColumn) {
            return switch (column) {
                case 0 -> "Блок-секция";
                case 1 -> "Этаж";
                case 2 -> "Помещение";
                case 3 -> "Комната";
                case 4 -> "Кол-во каналов";
                case 5 -> "Сечение (кв.м)";
                case 6 -> "Объем (куб.м)";
                default -> "";
            };
        } else {
            return switch (column) {
                case 0 -> "Этаж";
                case 1 -> "Помещение";
                case 2 -> "Комната";
                case 3 -> "Кол-во каналов";
                case 4 -> "Сечение (кв.м)";
                case 5 -> "Объем (куб.м)";
                default -> "";
            };
        }
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        if (showSectionColumn) {
            return switch (columnIndex) {
                case 0 -> Integer.class;        // Блок-секция
                case 1, 2, 3 -> String.class;
                case 4 -> Integer.class;        // каналы
                case 5, 6 -> Double.class;      // сечение, объем
                default -> Object.class;
            };
        } else {
            return switch (columnIndex) {
                case 0, 1, 2 -> String.class;
                case 3 -> Integer.class;        // каналы
                case 4, 5 -> Double.class;      // сечение, объем
                default -> Object.class;
            };
        }
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        // редактируем только кол-во каналов, сечение, объем
        if (showSectionColumn) return columnIndex == 4 || columnIndex == 5 || columnIndex == 6;
        return columnIndex == 3 || columnIndex == 4 || columnIndex == 5;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        VentilationRecord r = records.get(rowIndex);
        if (showSectionColumn) {
            return switch (columnIndex) {
                case 0 -> r.sectionIndex() + 1;   // показываем 1..N
                case 1 -> r.floor();
                case 2 -> r.space();
                case 3 -> r.room();
                case 4 -> r.channels();
                case 5 -> r.sectionArea();
                case 6 -> r.volume();
                default -> "";
            };
        } else {
            return switch (columnIndex) {
                case 0 -> r.floor();
                case 1 -> r.space();
                case 2 -> r.room();
                case 3 -> r.channels();
                case 4 -> r.sectionArea();
                case 5 -> r.volume();
                default -> "";
            };
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        VentilationRecord r = records.get(rowIndex);
        try {
            if (showSectionColumn) {
                if (columnIndex == 4) {
                    int newVal = toInt(aValue, r.channels());
                    VentilationRecord nr = r.withChannels(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationChannels(newVal);
                } else if (columnIndex == 5) {
                    double newVal = toDouble(aValue, r.sectionArea());
                    VentilationRecord nr = r.withSectionArea(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(newVal);
                } else if (columnIndex == 6) {
                    Double newVal = toNullableDouble(aValue, r.volume());
                    VentilationRecord nr = r.withVolume(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVolume(newVal);
                } else return;
            } else {
                if (columnIndex == 3) {
                    int newVal = toInt(aValue, r.channels());
                    VentilationRecord nr = r.withChannels(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationChannels(newVal);
                } else if (columnIndex == 4) {
                    double newVal = toDouble(aValue, r.sectionArea());
                    VentilationRecord nr = r.withSectionArea(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVentilationSectionArea(newVal);
                } else if (columnIndex == 5) {
                    Double newVal = toNullableDouble(aValue, r.volume());
                    VentilationRecord nr = r.withVolume(newVal);
                    records.set(rowIndex, nr);
                    if (nr.roomRef() != null) nr.roomRef().setVolume(newVal);
                } else return;
            }
            fireTableRowsUpdated(rowIndex, rowIndex);
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
}
