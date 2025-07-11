package ru.citlab24.protokol.tabs.renderers;

import javax.swing.*;
import java.awt.*;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.tabs.models.Room;

public class RoomListRenderer extends DefaultListCellRenderer {
    @Override
    public Component getListCellRendererComponent(JList<?> list,
                                                  Object value,
                                                  int index,
                                                  boolean isSelected,
                                                  boolean cellHasFocus) {
        super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
        if (value instanceof Room) {
            Room room = (Room) value;
            setText(room.getName());
            setIcon(FontIcon.of(FontAwesomeSolid.CUBE, 16, new Color(156, 39, 176)));
        }
        return this;
    }
}
