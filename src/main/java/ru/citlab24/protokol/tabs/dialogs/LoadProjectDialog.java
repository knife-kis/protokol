package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.db.DatabaseManager;

import javax.swing.*;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.sql.SQLException;
import java.util.List;

public class LoadProjectDialog extends JDialog {
    private final DefaultListModel<Building> model = new DefaultListModel<>();
    private final JList<Building> list = new JList<>(model);
    private Building selectedProject;

    public LoadProjectDialog(JFrame parent, List<Building> projects) {
        super(parent, "Загрузить проект", true);
        setSize(520, 420);
        setLocationRelativeTo(parent);
        setLayout(new BorderLayout(10, 10));

        // Список проектов
        list.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        list.setFixedCellHeight(28);
        list.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index,
                                                          boolean isSelected, boolean cellHasFocus) {
                Component c = super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Building) {
                    setText(((Building) value).getName());
                }
                return c;
            }
        });

        for (Building p : projects) model.addElement(p);

        // Двойной клик = ОК
        list.addMouseListener(new MouseAdapter() {
            @Override public void mouseClicked(MouseEvent e) {
                if (e.getClickCount() == 2 && !list.isSelectionEmpty()) {
                    onOk();
                }
            }
        });

        add(new JScrollPane(list), BorderLayout.CENTER);

        // Кнопки: Удалить / ОК / Отмена
        JButton deleteBtn = new JButton("Удалить…");
        JButton okBtn     = new JButton("ОК");
        JButton cancelBtn = new JButton("Отмена");

        deleteBtn.addActionListener(e -> onDelete());
        okBtn.addActionListener(e -> onOk());
        cancelBtn.addActionListener(e -> dispose());

        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btns.add(deleteBtn);
        btns.add(okBtn);
        btns.add(cancelBtn);
        add(btns, BorderLayout.SOUTH);

        // Состояние кнопок
        list.addListSelectionListener(e -> {
            boolean hasSel = !list.isSelectionEmpty();
            deleteBtn.setEnabled(hasSel);
            okBtn.setEnabled(hasSel);
        });
        boolean hasSel = !list.isSelectionEmpty();
        deleteBtn.setEnabled(hasSel);
        okBtn.setEnabled(hasSel);
    }

    private void onOk() {
        if (!list.isSelectionEmpty()) {
            selectedProject = list.getSelectedValue();
            dispose();
        }
    }

    private void onDelete() {
        if (list.isSelectionEmpty()) return;
        Building b = list.getSelectedValue();
        String name = (b.getName() != null ? b.getName() : ("ID " + b.getId()));

        int ans = JOptionPane.showConfirmDialog(
                this,
                "Вы точно хотите удалить проект:\n«" + name + "»?\nЭто действие необратимо.",
                "Подтвердите удаление",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.WARNING_MESSAGE
        );
        if (ans != JOptionPane.YES_OPTION) return;

        try {
            DatabaseManager.deleteBuilding(b.getId());
            // Убираем удалённый элемент из списка
            int idx = list.getSelectedIndex();
            model.remove(idx);
            if (!model.isEmpty()) {
                list.setSelectedIndex(Math.min(idx, model.getSize() - 1));
            }
            // Сообщение об успехе (по желанию)
            JOptionPane.showMessageDialog(this, "Проект удалён.", "Готово", JOptionPane.INFORMATION_MESSAGE);
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this, "Ошибка удаления: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }

    public Building getSelectedProject() {
        return selectedProject;
    }
}
