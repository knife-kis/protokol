package ru.citlab24.protokol.ui;

import javax.swing.*;
import java.awt.*;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Line2D;

/**
 * Кастомная «крестик-звезда» иконка для кнопки закрытия окна в title bar.
 */
public class TitleBarCloseIcon implements Icon {

    private static final int SIZE = 16;

    @Override
    public void paintIcon(Component c, Graphics g, int x, int y) {
        Graphics2D g2 = (Graphics2D) g.create();
        try {
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setRenderingHint(RenderingHints.KEY_STROKE_CONTROL, RenderingHints.VALUE_STROKE_PURE);

            // Чуть мягче обычного крестика: скруглённые линии + центральная точка.
            g2.setColor(new Color(0xD64D63));
            g2.setStroke(new BasicStroke(1.9f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));

            double left = x + 3.5;
            double right = x + SIZE - 3.5;
            double top = y + 3.5;
            double bottom = y + SIZE - 3.5;
            g2.draw(new Line2D.Double(left, top, right, bottom));
            g2.draw(new Line2D.Double(right, top, left, bottom));

            g2.fill(new Ellipse2D.Double(x + SIZE / 2.0 - 1.15, y + SIZE / 2.0 - 1.15, 2.3, 2.3));
        } finally {
            g2.dispose();
        }
    }

    @Override
    public int getIconWidth() {
        return SIZE;
    }

    @Override
    public int getIconHeight() {
        return SIZE;
    }
}
