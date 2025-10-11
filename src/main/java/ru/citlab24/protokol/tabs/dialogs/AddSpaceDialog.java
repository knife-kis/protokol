package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.*;
import java.awt.*;

public class AddSpaceDialog extends JDialog {
    private boolean confirmed = false;
    private JTextField identifierField;
    private JComboBox<Space.SpaceType> spaceTypeCombo;

    public AddSpaceDialog(JFrame parent, Floor.FloorType floorType) {
        super(parent, "Добавить помещение", true);
        setSize(350, 200);
        setLocationRelativeTo(parent);
        initUI(floorType);
    }

    private void initUI(Floor.FloorType floorType) {
        JPanel mainPanel = new JPanel(new GridLayout(3, 2, 10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(15, 15, 15, 15));

        // Поля ввода
        mainPanel.add(new JLabel("Идентификатор:"));
        identifierField = new JTextField();
        mainPanel.add(identifierField);

        mainPanel.add(new JLabel("Тип помещения:"));
        spaceTypeCombo = new JComboBox<>(getAllowedSpaceTypes(floorType));
        mainPanel.add(spaceTypeCombo);

        // Кнопки
        JButton okButton = new JButton("Добавить");
        okButton.addActionListener(e -> {
            confirmed = true;
            dispose();
        });

        JButton cancelButton = new JButton("Отмена");
        cancelButton.addActionListener(e -> dispose());

        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        buttonPanel.add(okButton);
        buttonPanel.add(cancelButton);

        add(mainPanel, BorderLayout.CENTER);
        add(buttonPanel, BorderLayout.SOUTH);
    }

    private Space.SpaceType[] getAllowedSpaceTypes(Floor.FloorType floorType) {
        if (floorType == null) return Space.SpaceType.values();
        switch (floorType) {
            case RESIDENTIAL:
                return new Space.SpaceType[] { Space.SpaceType.APARTMENT };
            case OFFICE:
                return new Space.SpaceType[] { Space.SpaceType.OFFICE };
            case PUBLIC:
                return Space.SpaceType.values();   // общественные — разрешаем любые
            case MIXED:
                return Space.SpaceType.values();   // смешанные — разрешаем любые
            default:
                return Space.SpaceType.values();
        }
    }

    public boolean showDialog() {
        setVisible(true);
        return confirmed;
    }

    public String getSpaceIdentifier() {
        return identifierField.getText().trim();
    }

    public Space.SpaceType getSpaceType() {
        return (Space.SpaceType) spaceTypeCombo.getSelectedItem();
    }
    public void setSpaceIdentifier(String identifier) {
        identifierField.setText(identifier);
    }

    public void setSpaceType(Space.SpaceType type) {
        spaceTypeCombo.setSelectedItem(type);
    }
}
