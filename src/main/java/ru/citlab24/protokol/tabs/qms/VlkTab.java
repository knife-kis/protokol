package ru.citlab24.protokol.tabs.qms;

import javax.swing.*;
import java.awt.*;

public class VlkTab extends JPanel {

    private final JButton createPlanButton = new JButton("Создать план мониторинга достоверности результатов лабораторной деятельности");
    private final JLabel selectedYearLabel = new JLabel("Год плана не выбран", SwingConstants.CENTER);

    public VlkTab() {
        super(new BorderLayout(12, 12));

        createPlanButton.setFont(createPlanButton.getFont().deriveFont(Font.BOLD, 14f));
        createPlanButton.addActionListener(e -> requestPlanYear());

        selectedYearLabel.setFont(selectedYearLabel.getFont().deriveFont(Font.PLAIN, 16f));

        JPanel centerPanel = new JPanel(new GridBagLayout());
        centerPanel.add(createPlanButton);

        add(centerPanel, BorderLayout.CENTER);
        add(selectedYearLabel, BorderLayout.SOUTH);
    }

    private void requestPlanYear() {
        String input = JOptionPane.showInputDialog(
                this,
                "Введите год (4 цифры):",
                "Создание плана мониторинга",
                JOptionPane.QUESTION_MESSAGE
        );

        if (input == null) {
            return;
        }

        String year = input.trim();
        if (!year.matches("\\d{4}")) {
            JOptionPane.showMessageDialog(
                    this,
                    "Год должен содержать ровно 4 цифры.",
                    "Некорректный ввод",
                    JOptionPane.WARNING_MESSAGE
            );
            return;
        }

        selectedYearLabel.setText("Выбран год плана мониторинга: " + year);
    }
}
