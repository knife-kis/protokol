package ru.citlab24.protokol.tabs.modules.microclimateTab;

import javax.swing.*;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableCellRenderer;
import java.awt.*;

final class WallsButtonsCell extends AbstractCellEditor
        implements TableCellEditor, TableCellRenderer {

    private final JPanel panel;
    private final ButtonGroup group;
    private final JToggleButton[] btn = new JToggleButton[5];

    WallsButtonsCell() {
        panel = new JPanel(new FlowLayout(FlowLayout.CENTER, 4, 2));
        group = new ButtonGroup();
        for (int i = 0; i <= 4; i++) {
            btn[i] = new JToggleButton(String.valueOf(i));
            btn[i].setMargin(new Insets(1, 6, 1, 6));
            btn[i].addActionListener(e -> stopCellEditing()); // выбрать — и сразу подтвердить
            group.add(btn[i]);
            panel.add(btn[i]);
        }
        panel.setOpaque(true);
    }

    private void applyValue(Integer v) {
        group.clearSelection();
        if (v != null && v >= 0 && v <= 4) {
            btn[v].setSelected(true);
        }
    }

    private Integer readValue() {
        for (int i = 0; i <= 4; i++) {
            if (btn[i].isSelected()) return i;
        }
        return null; // ни одна не выбрана
    }

    @Override public Object getCellEditorValue() {
        return readValue();
    }

    @Override
    public Component getTableCellEditorComponent(JTable table, Object value,
                                                 boolean isSelected, int row, int column) {
        applyValue((Integer) value);
        panel.setBackground(table.getSelectionBackground());
        return panel;
    }

    @Override
    public Component getTableCellRendererComponent(JTable table, Object value,
                                                   boolean isSelected, boolean hasFocus,
                                                   int row, int column) {
        applyValue((Integer) value);
        panel.setBackground(isSelected ? table.getSelectionBackground() : table.getBackground());
        return panel;
    }
}
