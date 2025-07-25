package ru.citlab24.protokol.tabs.modules.med;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;

public class RadiationTableModel extends AbstractTableModel {
    private final String[] columnNames = {
            "№ п/п", "Помещение", "Результат измерения, мкЗв/ч", "Допустимый уровень, мкЗв/ч"
    };

    private List<RadiationRecord> records = new ArrayList<>();

    @Override
    public int getRowCount() {
        return records.size();
    }

    @Override
    public int getColumnCount() {
        return columnNames.length;
    }

    @Override
    public String getColumnName(int column) {
        return columnNames[column];
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        RadiationRecord record = records.get(rowIndex);
        switch (columnIndex) {
            case 0: return rowIndex + 1;
            case 1: return record.getRoom();
            case 2: return record.getMeasurementResult();
            case 3: return record.getPermissibleLevel();
            default: return null;
        }
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        RadiationRecord record = records.get(rowIndex);
        switch (columnIndex) {
            case 2:
                if (aValue instanceof Double) {
                    record.setMeasurementResult((Double) aValue);
                } else if (aValue instanceof String) {
                    try {
                        record.setMeasurementResult(Double.parseDouble((String) aValue));
                    } catch (NumberFormatException e) {
                        // Обработка ошибки
                    }
                }
                break;
            case 3:
                if (aValue instanceof Double) {
                    record.setPermissibleLevel((Double) aValue);
                } else if (aValue instanceof String) {
                    try {
                        record.setPermissibleLevel(Double.parseDouble((String) aValue));
                    } catch (NumberFormatException e) {
                        // Обработка ошибки
                    }
                }
                break;
        }
        fireTableCellUpdated(rowIndex, columnIndex);
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        // Разрешаем редактировать только столбцы с измерениями
        return columnIndex == 2 || columnIndex == 3;
    }

    public void addRecord(RadiationRecord record) {
        records.add(record);
        fireTableRowsInserted(records.size() - 1, records.size() - 1);
    }

    public void removeRecord(int rowIndex) {
        records.remove(rowIndex);
        fireTableRowsDeleted(rowIndex, rowIndex);
    }

    public List<RadiationRecord> getRecords() {
        return new ArrayList<>(records);
    }
}