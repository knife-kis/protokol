package ru.citlab24.protokol.tabs;

import javax.swing.*;
import javax.swing.table.TableCellEditor;
import java.awt.*;

public class SpinnerEditor extends DefaultCellEditor {
    private final JSpinner spinner;

    public SpinnerEditor(double min, double max, double step) {
        this(min, min, max, step);
    }

    public SpinnerEditor(double value, double min, double max, double step) {
        super(new JTextField());
        SpinnerNumberModel model = new SpinnerNumberModel(value, min, max, step);
        spinner = new JSpinner(model);
        spinner.addChangeListener(e -> fireEditingStopped());
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value,
                                                 boolean isSelected, int row, int column) {
        spinner.setValue(value != null ? value : 0.0);
        return spinner;
    }

    @Override
    public Object getCellEditorValue() {
        Object value = spinner.getValue();
        return (value instanceof Double && (Double) value == 0.0) ? null : value;
    }
}