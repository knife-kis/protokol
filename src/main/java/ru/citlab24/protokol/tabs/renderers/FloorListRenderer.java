package ru.citlab24.protokol.tabs.renderers;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.tabs.models.Floor;

public class FloorListRenderer extends DefaultListCellRenderer {
    @Override
    public Component getListCellRendererComponent(JList<?> list, Object value, int index,
                                                  boolean isSelected, boolean cellHasFocus) {
        super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
        if (value instanceof Floor) {
            Floor floor = (Floor) value;
            // Используем getNumber() вместо getName()
            setText(String.format("Этаж %s: %s", floor.getNumber(), floor.getType()));
            setIcon(FontIcon.of(FontAwesomeSolid.BUILDING, 16, new Color(0, 115, 200)));
        }
        return this;
    }
}