package ru.citlab24.protokol.tabs;

import javax.swing.*;
import javax.swing.table.TableCellEditor;
import java.awt.*;
import java.awt.event.MouseEvent;
import java.util.EventObject;

public class SpinnerEditor extends AbstractCellEditor implements TableCellEditor {
    private final JSpinner spinner;
    private final SpinnerNumberModel model;
    private Double originalValue;
    private final double minValue;

    public SpinnerEditor(double initial, double min, double max, double step) {
        model = new SpinnerNumberModel(initial, min, max, step);
        spinner = new JSpinner(model);
        spinner.setBorder(BorderFactory.createEmptyBorder());
        minValue = min;
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value,
                                                 boolean isSelected, int row, int column) {
        if (value == null) {
            originalValue = null;
            spinner.setValue(minValue);
        } else {
            originalValue = (value instanceof Number) ? ((Number) value).doubleValue() : 0.0;
            spinner.setValue(originalValue);
        }
        return spinner;
    }

    @Override
    public Object getCellEditorValue() {
        Object value = spinner.getValue();
        if (value instanceof Number) {
            double newValue = ((Number) value).doubleValue();

            // Возвращаем null если значение сброшено до минимума
            if (originalValue == null && Math.abs(newValue - minValue) < 0.000001) {
                return null;
            }
            return newValue;
        }
        return originalValue;
    }

    @Override
    public boolean isCellEditable(EventObject evt) {
        if (evt instanceof MouseEvent) {
            return ((MouseEvent) evt).getClickCount() >= 1;
        }
        return true;
    }
}