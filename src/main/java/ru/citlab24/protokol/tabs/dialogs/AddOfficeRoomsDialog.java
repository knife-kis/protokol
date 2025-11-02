package ru.citlab24.protokol.tabs.dialogs;

import javax.swing.*;
import java.awt.*;
import java.awt.event.*;
import java.util.*;

/**
 * Диалог пакетного добавления/редактирования комнат ДЛЯ ОФИСНЫХ помещений.
 * — Снизу широкая строка ввода.
 * — Справа «Быстрый выбор (офис)»; пополняется мгновенно при добавлении.
 * — По центру список «Будут добавлены» с перемещением/переименованием/удалением.
 * — Enter: добавить в список; Ctrl+Enter: OK; Esc: Отмена.
 * — Ctrl+клик по быстрым — сразу в список.
 * — Выравнивание шапок слева/справа; разделитель стартует на 70% слева; компактные отступы.
 */
public class AddOfficeRoomsDialog extends JDialog {

    // ===== ввод =====
    private final JTextField input = new JTextField();

    // ===== список по центру =====
    private final DefaultListModel<String> model = new DefaultListModel<>();
    private final JList<String> list = new JList<>(model);

    // Список «новых в этой сессии» — выделяем полужирным
    private final Set<String> sessionNew = new HashSet<>();

    // ===== быстрый выбор справа =====
    private final JPanel quickRight = new JPanel();
    private final java.util.Deque<String> recentQuick = new ArrayDeque<>();
    private static final int QUICK_LIMIT = 30;

    // Базовые подсказки, пришедшие извне (из BuildingTab)
    private final java.util.List<String> baseSuggestions;

    // Контроль уникальности в текущем наборе
    private final Set<String> unique = new HashSet<>();
    private boolean confirmed = false;
    private final boolean editMode;

    // флаг для Ctrl+клика
    private boolean ctrlClickDown = false;

    public AddOfficeRoomsDialog(JFrame parent,
                                java.util.List<String> suggestions,
                                String prefillName,
                                boolean editMode) {
        super(parent, editMode ? "Редактировать (офис)" : "Добавить комнаты (офис)", true);
        this.baseSuggestions = (suggestions != null && !suggestions.isEmpty())
                ? suggestions
                : defaultOfficeSuggestions();
        this.editMode = editMode;

        // окно +20%
        Dimension pref = new Dimension(864, 528);
        setPreferredSize(pref);
        setMinimumSize(pref);

        setLayout(new BorderLayout(10, 10));
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);

        add(buildCenter(), BorderLayout.CENTER);
        add(buildBottomBar(), BorderLayout.SOUTH);
        add(buildButtons(), BorderLayout.NORTH);

        if (editMode && prefillName != null && !prefillName.isBlank()) {
            addToModel(clean(prefillName));
        }

        // клавиши
        getRootPane().registerKeyboardAction(
                e -> { confirmed = false; setVisible(false); dispose(); },
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW);

        getRootPane().registerKeyboardAction(
                e -> onOk(),
                KeyStroke.getKeyStroke(KeyEvent.VK_ENTER, KeyEvent.CTRL_DOWN_MASK),
                JComponent.WHEN_IN_FOCUSED_WINDOW);

        input.addActionListener(e -> addCurrentText());

        pack();
        setLocationRelativeTo(parent);
    }

    private JComponent buildCenter() {
        // слева — список с тулбаром
        JPanel left = new JPanel(new BorderLayout(6, 6));
        left.add(buildListToolbar(), BorderLayout.NORTH);

        list.setVisibleRowCount(12);
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        list.setCellRenderer(new FancyListRenderer());
        left.add(new JScrollPane(list), BorderLayout.CENTER);

        list.addMouseListener(new MouseAdapter() {
            @Override public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2) renameSelected();
            }
        });

        // справа — быстрый выбор
        quickRight.setLayout(new BoxLayout(quickRight, BoxLayout.Y_AXIS));
        quickRight.setBorder(BorderFactory.createEmptyBorder(6,6,6,6));
        fillQuickButtons();

        JScrollPane rightScroll = new JScrollPane(quickRight);
        rightScroll.setPreferredSize(new Dimension(260, 200));

        JPanel rightWrap = new JPanel(new BorderLayout());
        JPanel rightHeader = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 6));
        rightHeader.add(new JLabel("Быстрый выбор (офис)"));
        rightWrap.add(rightHeader, BorderLayout.NORTH);
        rightWrap.add(rightScroll, BorderLayout.CENTER);

        // разделитель
        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT, left, rightWrap);
        split.setContinuousLayout(true);
        javax.swing.SwingUtilities.invokeLater(() -> split.setDividerLocation(1.80));
        split.setResizeWeight(0.5);
        return split;
    }

    private JComponent buildListToolbar() {
        JToolBar tb = new JToolBar();
        tb.setFloatable(false);

        JButton up = new JButton("↑"); up.setToolTipText("Выше"); up.addActionListener(e -> moveSelected(-1));
        JButton dn = new JButton("↓"); dn.setToolTipText("Ниже"); dn.addActionListener(e -> moveSelected(+1));
        JButton rn = new JButton("Переименовать"); rn.addActionListener(e -> renameSelected());
        JButton rm = new JButton("Удалить"); rm.addActionListener(e -> removeSelected());
        JButton cl = new JButton("Очистить");
        cl.addActionListener(e -> { model.clear(); unique.clear(); sessionNew.clear(); input.requestFocusInWindow(); });

        tb.add(up); tb.add(dn);
        tb.addSeparator();
        tb.add(rn); tb.add(rm);
        tb.addSeparator();
        tb.add(cl);
        return tb;
    }

    private JComponent buildBottomBar() {
        JPanel south = new JPanel(new BorderLayout(6, 4));
        south.setBorder(BorderFactory.createEmptyBorder(2, 8, 4, 8));

        input.setColumns(40);
        south.add(input, BorderLayout.CENTER);

        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT, 6, 2));
        JButton add = new JButton("Добавить");
        add.setToolTipText("Enter");
        add.addActionListener(e -> addCurrentText());
        btns.add(add);

        south.add(btns, BorderLayout.EAST);
        return south;
    }

    private JComponent buildButtons() {
        JPanel north = new JPanel(new FlowLayout(FlowLayout.RIGHT, 6, 2));
        north.setBorder(BorderFactory.createEmptyBorder(2, 8, 6, 8));

        JButton cancel = new JButton("Отмена");
        JButton ok = new JButton(editMode ? "Сохранить" : "Добавить комнаты");
        cancel.addActionListener(e -> { confirmed = false; setVisible(false); dispose(); });
        ok.addActionListener(e -> onOk());

        north.add(cancel);
        north.add(ok);
        getRootPane().setDefaultButton(ok);
        return north;
    }

    private void fillQuickButtons() {
        quickRight.removeAll();

        JLabel tip = new JLabel("<html><small>Клик — в поле ввода; Ctrl+клик — сразу в список.</small></html>");
        tip.setAlignmentX(Component.LEFT_ALIGNMENT);
        quickRight.add(tip);
        quickRight.add(Box.createVerticalStrut(6));

        // пул: recent → base
        LinkedHashSet<String> pool = new LinkedHashSet<>();
        for (String r : recentQuick) pool.add(r);
        for (String b : baseSuggestions) if (b != null && !b.isBlank()) pool.add(b);

        int i = 0;
        for (String s : pool) {
            if (i >= QUICK_LIMIT) break;
            JButton b = makeQuickButton(s);
            quickRight.add(b);
            quickRight.add(Box.createVerticalStrut(4));
            i++;
        }
        quickRight.revalidate();
        quickRight.repaint();
    }

    private JButton makeQuickButton(String text) {
        JButton b = new JButton(text);
        b.setAlignmentX(Component.LEFT_ALIGNMENT);
        b.setFocusPainted(false);

        b.addMouseListener(new MouseAdapter() {
            @Override public void mousePressed(MouseEvent me) {
                ctrlClickDown =
                        me.isControlDown()
                                || ((me.getModifiersEx() & InputEvent.CTRL_DOWN_MASK) != 0)
                                || ((me.getModifiersEx() & InputEvent.META_DOWN_MASK) != 0);
            }
        });

        b.addActionListener(e -> {
            if (ctrlClickDown) {
                addToModel(clean(text));
            } else {
                input.setText(text);
                input.requestFocusInWindow();
                input.selectAll();
            }
            ctrlClickDown = false;
        });

        return b;
    }

    private void addCurrentText() {
        String txt = clean(input.getText());
        if (txt.isBlank()) {
            JOptionPane.showMessageDialog(this, "Введите название комнаты", "Пустое значение", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (addToModel(txt)) {
            input.setText("");
            input.requestFocusInWindow();
        }
    }

    private boolean addToModel(String name) {
        String key = normalize(name);
        if (key.isEmpty()) return false;
        if (unique.contains(key)) {
            int idx = indexOfKey(key);
            if (idx >= 0) { list.setSelectedIndex(idx); list.ensureIndexIsVisible(idx); }
            return false;
        }
        unique.add(key);
        model.addElement(name);
        sessionNew.add(key);
        list.setSelectedIndex(model.size() - 1);
        list.ensureIndexIsVisible(model.getSize() - 1);

        addToRecentQuick(name);
        fillQuickButtons();
        return true;
    }

    private void addToRecentQuick(String name) {
        String s = clean(name);
        recentQuick.removeIf(x -> x.equalsIgnoreCase(s));
        recentQuick.addFirst(s);
        while (recentQuick.size() > QUICK_LIMIT) recentQuick.removeLast();
    }

    private int indexOfKey(String key) {
        for (int i = 0; i < model.size(); i++) if (normalize(model.get(i)).equals(key)) return i;
        return -1;
    }

    private void moveSelected(int d) {
        int i = list.getSelectedIndex(); if (i < 0) return;
        int j = i + d; if (j < 0 || j >= model.size()) return;
        String a = model.get(i);
        model.set(i, model.get(j));
        model.set(j, a);
        list.setSelectedIndex(j);
        list.ensureIndexIsVisible(j);
    }

    private void renameSelected() {
        int i = list.getSelectedIndex(); if (i < 0) return;
        String old = model.get(i);
        String neu = prompt("Переименовать", "Новое название:", old);
        if (neu == null) return;
        neu = clean(neu);
        String key = normalize(neu);
        if (key.isEmpty()) return;

        if (unique.contains(key) && !normalize(old).equals(key)) {
            int idx = indexOfKey(key);
            if (idx >= 0) {
                list.setSelectedIndex(idx);
                list.ensureIndexIsVisible(idx);
                model.remove(i);
                unique.remove(normalize(old));
                sessionNew.remove(normalize(old));
            }
            return;
        }
        unique.remove(normalize(old));
        sessionNew.remove(normalize(old));
        unique.add(key);
        sessionNew.add(key);
        model.set(i, neu);

        addToRecentQuick(neu);
        fillQuickButtons();
    }

    private void removeSelected() {
        int i = list.getSelectedIndex(); if (i < 0) return;
        String key = normalize(model.get(i));
        unique.remove(key);
        sessionNew.remove(key);
        model.remove(i);
        if (!model.isEmpty()) list.setSelectedIndex(Math.min(i, model.size() - 1));
    }

    private void onOk() {
        if (editMode) {
            if (model.isEmpty()) {
                String v = clean(input.getText());
                if (v.isBlank()) {
                    JOptionPane.showMessageDialog(this, "Введите название комнаты", "Пустое значение", JOptionPane.WARNING_MESSAGE);
                    return;
                }
                addToModel(v);
            }
        } else {
            if (model.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Добавьте хотя бы одну комнату", "Пустой список", JOptionPane.WARNING_MESSAGE);
                return;
            }
        }
        confirmed = true;
        setVisible(false);
        dispose();
    }

    private static String clean(String s) {
        if (s == null) return "";
        String t = s.replace('\u00A0',' ').trim();
        return t.replaceAll("\\s+", " ");
    }
    private static String normalize(String s) { return clean(s).toLowerCase(java.util.Locale.ROOT); }

    private String prompt(String title, String label, String init) {
        JTextField tf = new JTextField(init == null ? "" : init);
        JPanel p = new JPanel(new BorderLayout(8, 0));
        p.add(new JLabel(label), BorderLayout.WEST);
        p.add(tf, BorderLayout.CENTER);

        final JOptionPane pane = new JOptionPane(p, JOptionPane.PLAIN_MESSAGE, JOptionPane.OK_CANCEL_OPTION);
        final JDialog dialog = pane.createDialog(this, title);

        tf.addActionListener(e -> { pane.setValue(JOptionPane.OK_OPTION); dialog.dispose(); });
        dialog.getRootPane().registerKeyboardAction(
                e -> { pane.setValue(JOptionPane.CANCEL_OPTION); dialog.dispose(); },
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW);

        dialog.setVisible(true);
        Object v = pane.getValue();
        boolean ok = (v != null) && Integer.valueOf(JOptionPane.OK_OPTION).equals(v);
        return ok ? tf.getText() : null;
    }

    // ==== публичный API ====
    public boolean showDialog() { setVisible(true); return confirmed; }

    public java.util.List<String> getNamesToAddList() {
        if (!confirmed) return java.util.List.of();
        java.util.List<String> out = new java.util.ArrayList<>();
        for (int i = 0; i < model.size(); i++) {
            String v = clean(model.get(i));
            if (!v.isBlank()) out.add(v);
        }
        return out;
    }

    public String getNameToAdd() { // для editMode
        if (!confirmed) return null;
        if (!model.isEmpty()) return clean(model.get(0));
        String v = clean(input.getText());
        return v.isBlank() ? null : v;
    }

    // ===== рендерер левого списка =====
    private final class FancyListRenderer extends DefaultListCellRenderer {
        private final Color selBg = new Color(25,118,210);
        private final Color selFg = Color.WHITE;
        private final Color normBg = UIManager.getColor("List.background");
        private final Color normFg = UIManager.getColor("List.foreground");

        @Override
        public Component getListCellRendererComponent(JList<?> list, Object value, int index, boolean isSelected, boolean cellHasFocus) {
            JLabel lbl = (JLabel) super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
            String text = (value == null) ? "" : value.toString();

            if (isSelected) { lbl.setBackground(selBg); lbl.setForeground(selFg); }
            else { lbl.setBackground(normBg); lbl.setForeground(normFg); }

            boolean isNew = sessionNew.contains(normalize(text));
            lbl.setFont(lbl.getFont().deriveFont(isNew ? Font.BOLD : Font.PLAIN));
            lbl.setBorder(BorderFactory.createEmptyBorder(4, 8, 4, 8));
            return lbl;
        }
    }

    // ===== дефолтные офисные подсказки на случай пустой базы =====
    private static java.util.List<String> defaultOfficeSuggestions() {
        return java.util.List.of(
                "Кабинет", "Оpen space", "Переговорная", "Ресепшн",
                "Серверная", "Архив", "Склад", "Кухня",
                "Комната отдыха", "Кладовая", "Гардероб", "Коридор", "Санузел"
        );
    }
}
