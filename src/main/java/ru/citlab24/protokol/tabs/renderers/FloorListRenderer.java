package ru.citlab24.protokol.tabs.renderers;

import javax.swing.*;
import javax.swing.border.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Space;

public class FloorListRenderer extends DefaultListCellRenderer {
    @Override
    public Component getListCellRendererComponent(JList<?> list, Object value, int index,
                                                  boolean isSelected, boolean cellHasFocus) {
        super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
        if (value instanceof Floor) {
            Floor floor = (Floor) value;
            String typeStr = (floor.getType() != null)
                    ? floor.getType().toString().toLowerCase()
                    : "неизвестный тип";
            setText(floor.getNumber() + " (" + typeStr + ")");
        }
        return this;
    }
}