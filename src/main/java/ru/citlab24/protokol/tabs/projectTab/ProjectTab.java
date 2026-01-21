package ru.citlab24.protokol.tabs.projectTab;

import javax.swing.*;
import java.awt.*;

public class ProjectTab extends JPanel {
    public ProjectTab() {
        initComponents();
    }

    private void initComponents() {
        setLayout(new BorderLayout());
        JLabel info = new JLabel("Действия по проекту доступны в верхней панели.");
        info.setBorder(BorderFactory.createEmptyBorder(16, 16, 16, 16));
        add(info, BorderLayout.NORTH);
    }
}
