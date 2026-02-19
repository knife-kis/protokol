package ru.citlab24.protokol.tabs.qms;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ShewhartMapTab extends JPanel {

    private static final String[] MONTH_NAMES = {
            "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
            "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"
    };

    private final List<JTextField> monthFileFields = new ArrayList<>();

    public ShewhartMapTab() {
        super(new BorderLayout(10, 10));
        add(createTopPanel(), BorderLayout.NORTH);
        add(createFilesPanel(), BorderLayout.CENTER);
        add(createBottomPanel(), BorderLayout.SOUTH);
    }

    private JComponent createTopPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        JLabel title = new JLabel("Карта Шухарта", SwingConstants.LEFT);
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));

        JLabel hint = new JLabel("Загрузите 12 Excel-файлов (по месяцам), затем нажмите «Преобразовать».", SwingConstants.LEFT);
        hint.setBorder(BorderFactory.createEmptyBorder(4, 0, 0, 0));

        panel.add(title, BorderLayout.NORTH);
        panel.add(hint, BorderLayout.SOUTH);
        panel.setBorder(BorderFactory.createEmptyBorder(8, 8, 0, 8));
        return panel;
    }

    private JComponent createFilesPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(BorderFactory.createTitledBorder("Входные файлы"));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 0;

        for (int i = 0; i < MONTH_NAMES.length; i++) {
            gbc.gridy = i;
            gbc.gridx = 0;
            panel.add(new JLabel(MONTH_NAMES[i] + ":"), gbc);

            JTextField pathField = new JTextField();
            pathField.setEditable(false);
            monthFileFields.add(pathField);

            gbc.gridx = 1;
            gbc.weightx = 1;
            panel.add(pathField, gbc);

            JButton chooseButton = new JButton("Выбрать");
            final int monthIndex = i;
            chooseButton.addActionListener(e -> chooseMonthFile(monthIndex));

            gbc.gridx = 2;
            gbc.weightx = 0;
            panel.add(chooseButton, gbc);
        }

        JScrollPane scrollPane = new JScrollPane(panel);
        scrollPane.getVerticalScrollBar().setUnitIncrement(20);
        scrollPane.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        return scrollPane;
    }

    private JComponent createBottomPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));

        JButton clearButton = new JButton("Очистить");
        clearButton.addActionListener(e -> monthFileFields.forEach(field -> field.setText("")));

        JButton convertButton = new JButton("Преобразовать");
        convertButton.setFont(convertButton.getFont().deriveFont(Font.BOLD, 14f));
        convertButton.addActionListener(e -> convertFiles());

        panel.add(clearButton);
        panel.add(convertButton);
        return panel;
    }

    private void chooseMonthFile(int monthIndex) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Выберите Excel для месяца: " + MONTH_NAMES[monthIndex]);
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xls", "xlsx", "xlsm"));

        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selected = chooser.getSelectedFile();
            monthFileFields.get(monthIndex).setText(selected.getAbsolutePath());
        }
    }

    private void convertFiles() {
        List<File> inputFiles = new ArrayList<>();
        for (JTextField monthFileField : monthFileFields) {
            String path = monthFileField.getText() == null ? "" : monthFileField.getText().trim();
            if (path.isEmpty()) {
                JOptionPane.showMessageDialog(
                        this,
                        "Необходимо выбрать все 12 файлов (по одному на каждый месяц).",
                        "Недостаточно файлов",
                        JOptionPane.WARNING_MESSAGE
                );
                return;
            }
            inputFiles.add(new File(path));
        }

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Сохранить результат в формате XLSM");
        chooser.setSelectedFile(new File("Карта_Шухарта.xlsm"));
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Macro-Enabled Workbook", "xlsm"));

        int result = chooser.showSaveDialog(this);
        if (result != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File targetFile = chooser.getSelectedFile();
        if (!targetFile.getName().toLowerCase().endsWith(".xlsm")) {
            targetFile = new File(targetFile.getParentFile(), targetFile.getName() + ".xlsm");
        }

        try {
            ShewhartMapExcelExporter.exportStaticTitle(targetFile, inputFiles);
            JOptionPane.showMessageDialog(
                    this,
                    "Файл успешно создан:\n" + targetFile.getAbsolutePath(),
                    "Готово",
                    JOptionPane.INFORMATION_MESSAGE
            );
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(
                    this,
                    "Не удалось создать файл:\n" + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }
}
