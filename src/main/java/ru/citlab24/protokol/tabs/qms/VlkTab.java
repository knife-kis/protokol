package ru.citlab24.protokol.tabs.qms;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.time.Year;
import java.util.ArrayList;
import java.util.List;

public class VlkTab extends JPanel {

    private static final String[] EXTRA_TABLE_COLUMNS = {
            "Наименование мероприятий",
            "Периодичность",
            "Ответственный",
            "Формы учета"
    };

    private final JTextField yearField = new JTextField(String.valueOf(Year.now().getValue()), 6);
    private final JButton createPlanButton = new JButton("Сформировать Word-файл");
    private final JButton addRowButton = new JButton("+ Добавить строку");
    private final JButton swapResponsiblesButton = new JButton("Поменять Тарновский ↔ Белов");
    private final JLabel selectedYearLabel = new JLabel("Год плана не выбран", SwingConstants.LEFT);

    private final DefaultTableModel extraRowsModel = new DefaultTableModel(EXTRA_TABLE_COLUMNS, 0) {
        @Override
        public boolean isCellEditable(int row, int column) {
            return true;
        }
    };

    public VlkTab() {
        super(new BorderLayout(10, 10));

        add(createToolbar(), BorderLayout.NORTH);
        add(createExtraRowsPanel(), BorderLayout.CENTER);
        add(selectedYearLabel, BorderLayout.SOUTH);

        fillDefaultExtraRows();
        wireActions();
    }

    private JComponent createToolbar() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.gridy = 0;

        JLabel yearLabel = new JLabel("Год:");
        yearLabel.setFont(yearLabel.getFont().deriveFont(Font.BOLD));

        createPlanButton.setFont(createPlanButton.getFont().deriveFont(Font.BOLD, 14f));

        gbc.gridx = 0;
        panel.add(yearLabel, gbc);

        gbc.gridx = 1;
        panel.add(yearField, gbc);

        gbc.gridx = 2;
        panel.add(createPlanButton, gbc);

        gbc.gridx = 3;
        panel.add(addRowButton, gbc);

        gbc.gridx = 4;
        panel.add(swapResponsiblesButton, gbc);

        return panel;
    }

    private JComponent createExtraRowsPanel() {
        JTable table = new JTable(extraRowsModel);
        table.setRowHeight(28);
        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);

        table.getColumnModel().getColumn(0).setPreferredWidth(520);
        table.getColumnModel().getColumn(1).setPreferredWidth(300);
        table.getColumnModel().getColumn(2).setPreferredWidth(180);
        table.getColumnModel().getColumn(3).setPreferredWidth(320);

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setBorder(BorderFactory.createTitledBorder("Дополнительные строки (редактируемые)") );
        scrollPane.getVerticalScrollBar().setUnitIncrement(24);
        return scrollPane;
    }

    private void wireActions() {
        addRowButton.addActionListener(e -> extraRowsModel.addRow(new String[]{"", "", "", ""}));
        swapResponsiblesButton.addActionListener(e -> swapResponsiblesInExtraRows());
        createPlanButton.addActionListener(e -> createPlanFile());
    }

    private void fillDefaultExtraRows() {
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МИ М.08-2021",
                "Четыре раза в год",
                "Тарновский М.О.",
                "Журнале регистрации результатов методом повторных исследований (испытаний) и измерений"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МИ РД.10-2021",
                "Четыре раза в год",
                "Тарновский М.О.",
                "Журнале регистрации результатов методом повторных исследований (испытаний) и измерений"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МИ Ш.13-2022",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МУ МР 2.6.1.0361-24",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МУ 2.6.1.0333-23",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по «Методика измерения плотности потока радона с поверхности земли и строительных конструкций (свидетельство об аттестации МВИ № 40090.6К816)»",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по «Паспорт (техническое описание, инструкция по эксплуатации, формуляр) на дозиметр-радиометр «ДРБП-03»",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по «Руководство по эксплуатации «Альфа-радиометр РАА-20П2»",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МУК 4.3.3722-21",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по Руководство по эксплуатации Экофизика-110А ПКДУ.411000.001.02 РЭ п.7.1, п.7.2 с приложением МИ ПКФ 12-006 п.2, п.5",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по ГОСТ 27296-2012, ГОСТ Р ИСО 3382-2-2013, ГОСТ Р 56769-2015, ГОСТ Р 56770-2015",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по МИ СС.09-2021",
                "Один раз в год",
                "Тарновский М.О.",
                "Протокол по результатам наблюдения"
        });
    }

    private void swapResponsiblesInExtraRows() {
        for (int row = 0; row < extraRowsModel.getRowCount(); row++) {
            String value = String.valueOf(extraRowsModel.getValueAt(row, 2));
            if (value.contains("Тарновский М.О.")) {
                extraRowsModel.setValueAt(value.replace("Тарновский М.О.", "Белов Д.А."), row, 2);
            } else if (value.contains("Белов Д.А.")) {
                extraRowsModel.setValueAt(value.replace("Белов Д.А.", "Тарновский М.О."), row, 2);
            }
        }
    }

    private void createPlanFile() {
        String year = yearField.getText().trim();
        if (!year.matches("\\d{4}")) {
            JOptionPane.showMessageDialog(
                    this,
                    "Год должен содержать ровно 4 цифры.",
                    "Некорректный ввод",
                    JOptionPane.WARNING_MESSAGE
            );
            return;
        }

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Сохранить план мониторинга ВЛК");
        chooser.setSelectedFile(new File("План_мониторинга_ВЛК_" + year + ".docx"));

        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File file = chooser.getSelectedFile();
        if (!file.getName().toLowerCase().endsWith(".docx")) {
            file = new File(file.getParentFile(), file.getName() + ".docx");
        }

        try {
            VlkWordExporter.export(file, year, collectExtraRows());
            selectedYearLabel.setText("Сформирован файл: " + file.getAbsolutePath());
            JOptionPane.showMessageDialog(this, "Файл успешно создан:\n" + file.getAbsolutePath());
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(
                    this,
                    "Не удалось создать документ: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private List<VlkWordExporter.PlanRow> collectExtraRows() {
        List<VlkWordExporter.PlanRow> rows = new ArrayList<>();
        for (int i = 0; i < extraRowsModel.getRowCount(); i++) {
            String event = value(i, 0);
            String period = value(i, 1);
            String responsible = value(i, 2);
            String forms = value(i, 3);

            if (event.isBlank() && period.isBlank() && responsible.isBlank() && forms.isBlank()) {
                continue;
            }

            rows.add(new VlkWordExporter.PlanRow(event, period, responsible, forms, ""));
        }
        return rows;
    }

    private String value(int row, int col) {
        Object raw = extraRowsModel.getValueAt(row, col);
        return raw == null ? "" : raw.toString().trim();
    }
}
