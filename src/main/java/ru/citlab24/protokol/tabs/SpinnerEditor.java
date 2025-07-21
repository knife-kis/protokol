package ru.citlab24.protokol.tabs;

import javax.swing.*;
import javax.swing.table.TableCellEditor;
import java.awt.*;
import java.awt.event.MouseEvent;
import java.util.EventObject;

public class SpinnerEditor extends AbstractCellEditor implements TableCellEditor {
    private final JSpinner spinner;

    // Конструктор для целых чисел
    public SpinnerEditor(int min, int max, int step) {
        spinner = new JSpinner(new SpinnerNumberModel(min, min, max, step));
        initSpinner();
    }

    // Конструктор для дробных чисел
    public SpinnerEditor(double min, double max, double step) {
        spinner = new JSpinner(new SpinnerNumberModel(min, min, max, step));
        initSpinner();
    }

    // Общий конструктор с шагом
    public SpinnerEditor(double min, double max, double step, double initial) {
        spinner = new JSpinner(new SpinnerNumberModel(min, max, step, initial));
        initSpinner();
    }

    private void initSpinner() {
        // Завершаем редактирование при любом изменении значения
        spinner.addChangeListener(e -> fireEditingStopped());

        // Настраиваем редактор для правильного отображения чисел
        JSpinner.NumberEditor editor = new JSpinner.NumberEditor(spinner);
        spinner.setEditor(editor);
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value,
                                                 boolean isSelected,
                                                 int row, int column) {
        spinner.setValue(value);
        return spinner;
    }

    @Override
    public Object getCellEditorValue() {
        return spinner.getValue();
    }

    @Override
    public boolean isCellEditable(EventObject evt) {
        if (evt instanceof MouseEvent) {
            return ((MouseEvent)evt).getClickCount() >= 1;
        }
        return true;
    }
}