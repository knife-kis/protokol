package ru.citlab24.protokol.tabs.dialogs;

import javax.swing.*;
import java.awt.*;
import java.awt.event.KeyEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.util.List;
import java.util.Objects;

/**
 * Диалог добавления/редактирования комнаты:
 * - только поле ввода + «быстрый выбор» кнопками;
 * - НЕТ выпадающих подсказок, НЕТ автонумерации и количества.
 * Клик по кнопке «быстрого выбора» просто вставляет текст в поле.
 */
public class AddRoomDialog extends JDialog {

    private final JTextField nameField = new JTextField();
    private final JPanel quickPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 6));

    private boolean confirmed = false;

    private final List<String> suggestions; // популярные названия
    private final boolean editMode;

    public AddRoomDialog(JFrame parent,
                         List<String> suggestions,
                         String prefillName,
                         boolean editMode) {
        super(parent, editMode ? "Редактировать комнату" : "Добавить комнату", true);
        this.suggestions = (suggestions != null) ? suggestions : List.of();
        this.editMode = editMode;

        setMinimumSize(new Dimension(520, 220));
        setLayout(new BorderLayout(10, 10));
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);

        add(buildTopPanel(), BorderLayout.NORTH);
        add(buildCenterPanel(), BorderLayout.CENTER);
        add(buildButtons(), BorderLayout.SOUTH);

        // Стартовое значение: в режиме редактирования — префилл, в режиме добавления — ПУСТО
        if (prefillName != null && !prefillName.isBlank() && editMode) {
            nameField.setText(prefillName);
            nameField.selectAll();
        }

        // ESC = отмена
        getRootPane().registerKeyboardAction(
                e -> { confirmed = false; setVisible(false); dispose(); },
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW
        );

        pack();
        setLocationRelativeTo(parent);
        addWindowListener(new WindowAdapter() {
            @Override public void windowOpened(WindowEvent e) {
                nameField.requestFocusInWindow();
                // Если есть префилл — он уже selectAll(); если пусто, лишним не будет
                nameField.selectAll();
            }
        });

        // После показа — фокус в поле
        addWindowListener(new java.awt.event.WindowAdapter() {
            @Override public void windowOpened(java.awt.event.WindowEvent e) {
                nameField.requestFocusInWindow();
                if (editMode) nameField.selectAll();
            }
        });
    }

    private JPanel buildTopPanel() {
        JPanel p = new JPanel(new BorderLayout());
        JLabel lbl = new JLabel("Быстрый выбор:");
        p.add(lbl, BorderLayout.NORTH);

        quickPanel.setBorder(BorderFactory.createEmptyBorder(4, 4, 0, 4));
        fillQuickButtons();

        JScrollPane sp = new JScrollPane(
                quickPanel,
                ScrollPaneConstants.VERTICAL_SCROLLBAR_AS_NEEDED,
                ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER
        );
        p.add(sp, BorderLayout.CENTER);
        return p;
    }

    private void fillQuickButtons() {
        quickPanel.removeAll();
        int limit = Math.min(12, suggestions.size());
        for (int i = 0; i < limit; i++) {
            String s = suggestions.get(i);
            if (s == null || s.isBlank()) continue;
            JButton b = new JButton(s);
            b.setFocusPainted(false);
            b.putClientProperty("JButton.buttonType", "toolBarButton");
            b.addActionListener(e -> {
                nameField.setText(s);
                nameField.requestFocusInWindow();
                nameField.selectAll(); // чтобы можно было сразу заменить при желании
            });
            quickPanel.add(b);
        }
        quickPanel.revalidate();
        quickPanel.repaint();
    }

    private JPanel buildCenterPanel() {
        JPanel form = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(6, 8, 6, 8);
        c.fill = GridBagConstraints.HORIZONTAL;

        c.gridx = 0; c.gridy = 0; c.weightx = 0;
        form.add(new JLabel("Название комнаты:"), c);

        c.gridx = 1; c.gridy = 0; c.weightx = 1;
        nameField.setColumns(30);
        form.add(nameField, c);

        return form;
    }

    private JPanel buildButtons() {
        JPanel south = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        JButton ok = new JButton(editMode ? "Сохранить" : "Добавить");
        JButton cancel = new JButton("Отмена");
        ok.addActionListener(e -> onOk());
        cancel.addActionListener(e -> {
            confirmed = false;
            setVisible(false);
            dispose();
        });
        south.add(cancel);
        south.add(ok);
        getRootPane().setDefaultButton(ok); // Enter = OK
        return south;
    }

    private void onOk() {
        String name = getEditorText().trim();
        if (name.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Введите название комнаты", "Пустое значение", JOptionPane.WARNING_MESSAGE);
            return;
        }
        confirmed = true;
        setVisible(false);
        dispose();
    }

    private String getEditorText() {
        Object it = nameField.getText();
        return (it == null) ? "" : Objects.toString(it, "");
    }

    public boolean showDialog() {
        setVisible(true);
        return confirmed;
    }

    /** Возвращает одно итоговое имя (или null, если Отмена) */
    public String getNameToAdd() {
        return confirmed ? getEditorText().trim() : null;
    }
}
