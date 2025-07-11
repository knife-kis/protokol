package ru.citlab24.protokol.tabs.renderers;

import javax.swing.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.tabs.models.Space;

public class SpaceListRenderer extends DefaultListCellRenderer {
    @Override
    public Component getListCellRendererComponent(JList<?> list,
                                                  Object value,
                                                  int index,
                                                  boolean isSelected,
                                                  boolean cellHasFocus) {
        super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
        if (value instanceof Space) {
            Space space = (Space) value;
            setText(String.format("%s: %s", space.getType(), space.getIdentifier()));
            setIcon(FontIcon.of(FontAwesomeSolid.DOOR_OPEN, 16, new Color(76, 175, 80)));
        }
        return this;
    }
}