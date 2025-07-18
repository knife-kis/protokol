// Новый класс LoadProjectDialog.java
package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Building;
import javax.swing.*;
import java.awt.*;
import java.util.List;

public class LoadProjectDialog extends JDialog {
    private Building selectedProject;
    private final JList<Building> projectList;

    public LoadProjectDialog(JFrame parent, List<Building> projects) {
        super(parent, "Выберите проект", true);
        setLayout(new BorderLayout());
        setSize(400, 300);
        setLocationRelativeTo(parent);

        DefaultListModel<Building> model = new DefaultListModel<>();
        for (Building project : projects) {
            model.addElement(project);
        }

        projectList = new JList<>(model);
        projectList.setCellRenderer(new DefaultListCellRenderer() {
            @Override
            public Component getListCellRendererComponent(JList<?> list, Object value, int index,
                                                          boolean isSelected, boolean cellHasFocus) {
                super.getListCellRendererComponent(list, value, index, isSelected, cellHasFocus);
                if (value instanceof Building) {
                    setText(((Building) value).getName());
                }
                return this;
            }
        });

        JButton loadButton = new JButton("Загрузить");
        loadButton.addActionListener(e -> {
            selectedProject = projectList.getSelectedValue();
            dispose();
        });

        JButton cancelButton = new JButton("Отмена");
        cancelButton.addActionListener(e -> dispose());

        JPanel buttonPanel = new JPanel();
        buttonPanel.add(loadButton);
        buttonPanel.add(cancelButton);

        add(new JScrollPane(projectList), BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);
    }

    public Building getSelectedProject() {
        return selectedProject;
    }
}