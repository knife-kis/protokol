package ru.citlab24.protokol.tabs.resourceTab;

import javax.swing.*;
import java.awt.*;

public class ResourceTab extends JPanel {
    private final CardLayout cardLayout = new CardLayout();
    private final JPanel content = new JPanel(cardLayout);

    public ResourceTab() {
        super(new BorderLayout(8, 8));

        JComboBox<String> selector = new JComboBox<>(new String[]{"Персонал", "Оборудование"});
        selector.addActionListener(e -> {
            String selected = (String) selector.getSelectedItem();
            cardLayout.show(content, selected == null ? "Персонал" : selected);
        });

        JPanel top = new JPanel(new FlowLayout(FlowLayout.LEFT));
        top.add(new JLabel("Раздел:"));
        top.add(selector);

        content.add(new PersonnelTab(), "Персонал");
        content.add(createEquipmentPlaceholder(), "Оборудование");

        add(top, BorderLayout.NORTH);
        add(content, BorderLayout.CENTER);
        cardLayout.show(content, "Персонал");
    }

    private JComponent createEquipmentPlaceholder() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.add(new JLabel("Раздел оборудования будет добавлен позже"));
        return panel;
    }
}
