package ru.citlab24.protokol;

import com.formdev.flatlaf.FlatClientProperties;
import com.formdev.flatlaf.FlatIntelliJLaf;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JComponent;
import javax.swing.JList;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTable;
import javax.swing.JTabbedPane;
import javax.swing.JViewport;
import javax.swing.UIManager;
import javax.swing.plaf.ColorUIResource;
import javax.swing.plaf.FontUIResource;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Font;

final class AppTheme {
    static final Color INK = new Color(0x202124);
    static final Color INK_SOFT = new Color(0x3A4047);
    static final Color SURFACE = new Color(0xF6FAFD);
    static final Color PANEL = new Color(0xEAF3FA);
    static final Color CARD = new Color(0xFFFFFF);
    static final Color BORDER = new Color(0xC6D6E3);
    static final Color ACCENT = new Color(0x2F80ED);
    static final Color ACCENT_DARK = new Color(0x1D5FBD);
    static final Color MENU = new Color(0x222325);
    static final Color MENU_HOVER = new Color(0x2C2E31);

    private AppTheme() {
    }

    static void install() {
        System.setProperty("awt.useSystemAAFontSettings", "on");
        System.setProperty("swing.aatext", "true");
        System.setProperty("flatlaf.useWindowDecorations", "true");
        com.formdev.flatlaf.FlatLaf.setUseNativeWindowDecorations(true);

        try {
            UIManager.setLookAndFeel(new FlatIntelliJLaf());
        } catch (Exception ignore) {
        }

        installFont();
        installColors();
        installMetrics();
        com.formdev.flatlaf.FlatLaf.updateUI();
    }

    static void decorateWorkspace(JComponent root) {
        if (root == null) return;
        decorateComponentTree(root);
    }

    private static void installFont() {
        FontUIResource font = new FontUIResource(new Font("Segoe UI", Font.PLAIN, 14));
        java.util.Enumeration<?> keys = UIManager.getDefaults().keys();
        while (keys.hasMoreElements()) {
            Object key = keys.nextElement();
            Object value = UIManager.get(key);
            if (value instanceof FontUIResource) {
                UIManager.put(key, font);
            }
        }

        UIManager.put("Table.font", new FontUIResource(new Font("Segoe UI", Font.PLAIN, 13)));
        UIManager.put("TableHeader.font", new FontUIResource(new Font("Segoe UI", Font.BOLD, 13)));
        UIManager.put("TabbedPane.font", new FontUIResource(new Font("Segoe UI", Font.BOLD, 13)));
        UIManager.put("Menu.font", new FontUIResource(new Font("Segoe UI", Font.BOLD, 13)));
        UIManager.put("MenuItem.font", new FontUIResource(new Font("Segoe UI", Font.PLAIN, 13)));
    }

    private static void installColors() {
        putColor("Panel.background", PANEL);
        putColor("Viewport.background", PANEL);
        putColor("ScrollPane.background", PANEL);
        putColor("TabbedPane.background", PANEL);
        putColor("TabbedPane.contentAreaColor", PANEL);
        putColor("TabbedPane.selectedBackground", CARD);
        putColor("TabbedPane.hoverColor", new Color(0xDDEDF8));
        putColor("TabbedPane.underlineColor", ACCENT);
        putColor("TabbedPane.foreground", INK_SOFT);
        putColor("TabbedPane.selectedForeground", INK);

        putColor("Label.foreground", INK);
        putColor("Button.background", CARD);
        putColor("Button.foreground", INK);
        putColor("Button.hoverBackground", new Color(0xE5F1FC));
        putColor("Button.pressedBackground", new Color(0xCFE4F8));
        putColor("Button.default.background", ACCENT);
        putColor("Button.default.foreground", Color.WHITE);
        putColor("Button.default.hoverBackground", ACCENT_DARK);

        putColor("TextField.background", CARD);
        putColor("FormattedTextField.background", CARD);
        putColor("PasswordField.background", CARD);
        putColor("ComboBox.background", CARD);
        putColor("TextField.foreground", INK);
        putColor("Component.borderColor", BORDER);
        putColor("Component.focusColor", ACCENT);

        putColor("Table.background", CARD);
        putColor("Table.foreground", INK);
        putColor("Table.gridColor", new Color(0xD8E3EC));
        putColor("Table.selectionBackground", new Color(0xD7EAFE));
        putColor("Table.selectionForeground", INK);
        putColor("TableHeader.background", new Color(0xEEF5FA));
        putColor("TableHeader.foreground", INK);
        putColor("List.background", CARD);
        putColor("List.selectionBackground", new Color(0xD7EAFE));
        putColor("List.selectionForeground", INK);

        putColor("MenuBar.background", MENU);
        putColor("MenuBar.foreground", Color.WHITE);
        putColor("Menu.background", MENU);
        putColor("Menu.foreground", Color.WHITE);
        putColor("Menu.selectionBackground", MENU_HOVER);
        putColor("Menu.selectionForeground", Color.WHITE);
        putColor("MenuItem.selectionBackground", new Color(0xE5F1FC));
        putColor("MenuItem.selectionForeground", INK);
        putColor("PopupMenu.background", CARD);
    }

    private static void installMetrics() {
        UIManager.put("Component.arc", 10);
        UIManager.put("Button.arc", 10);
        UIManager.put("TextComponent.arc", 8);
        UIManager.put("ComboBox.arc", 8);
        UIManager.put("Component.focusWidth", 1);
        UIManager.put("Button.innerFocusWidth", 0);
        UIManager.put("ScrollBar.width", 14);
        UIManager.put("ScrollBar.thumbArc", 999);
        UIManager.put("TabbedPane.tabArc", 8);
        UIManager.put("TabbedPane.tabInsets", new javax.swing.plaf.InsetsUIResource(6, 14, 6, 14));
        UIManager.put("Table.rowHeight", 28);
        UIManager.put("MenuBar.border", BorderFactory.createEmptyBorder(5, 10, 5, 10));
    }

    private static void putColor(String key, Color color) {
        UIManager.put(key, new ColorUIResource(color));
    }

    private static void decorateComponentTree(Component component) {
        if (component instanceof JPanel panel) {
            panel.setBackground(PANEL);
        } else if (component instanceof JScrollPane scrollPane) {
            scrollPane.setBorder(BorderFactory.createLineBorder(BORDER));
            scrollPane.setBackground(PANEL);
            scrollPane.getVerticalScrollBar().setUnitIncrement(16);
            JViewport viewport = scrollPane.getViewport();
            if (viewport != null) {
                viewport.setBackground(CARD);
            }
        } else if (component instanceof JSplitPane splitPane) {
            splitPane.setBorder(BorderFactory.createEmptyBorder());
            splitPane.setDividerSize(7);
            splitPane.setBackground(PANEL);
        } else if (component instanceof JTable table) {
            table.setRowHeight(Math.max(table.getRowHeight(), 28));
            table.setShowGrid(true);
            table.setGridColor(new Color(0xD8E3EC));
            table.setIntercellSpacing(new java.awt.Dimension(1, 1));
            table.setSelectionBackground(new Color(0xD7EAFE));
            table.setSelectionForeground(INK);
            table.getTableHeader().setBackground(new Color(0xEEF5FA));
            table.getTableHeader().setForeground(INK);
            table.getTableHeader().setReorderingAllowed(false);
        } else if (component instanceof JList<?> list) {
            list.setBackground(CARD);
            list.setSelectionBackground(new Color(0xD7EAFE));
            list.setSelectionForeground(INK);
            list.setFixedCellHeight(Math.max(list.getFixedCellHeight(), 28));
        } else if (component instanceof JButton button) {
            button.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
            if (button.getClientProperty(FlatClientProperties.STYLE) == null) {
                button.putClientProperty(FlatClientProperties.STYLE,
                        "arc: 8; focusWidth: 1; innerFocusWidth: 0");
            }
        } else if (component instanceof JTabbedPane tabs) {
            tabs.setBackground(PANEL);
            tabs.putClientProperty(FlatClientProperties.TABBED_PANE_HAS_FULL_BORDER, Boolean.FALSE);
            tabs.putClientProperty(FlatClientProperties.STYLE,
                    "tabType: card; showTabSeparators: false; tabArc: 8; contentSeparatorHeight: 0");
        }

        if (component instanceof Container container) {
            for (Component child : container.getComponents()) {
                decorateComponentTree(child);
            }
        }
    }
}
