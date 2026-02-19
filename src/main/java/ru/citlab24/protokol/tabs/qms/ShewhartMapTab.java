package ru.citlab24.protokol.tabs.qms;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.VlkDateRecord;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.io.IOException;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.LinkedHashSet;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ShewhartMapTab extends JPanel {

    private static final DateTimeFormatter DB_DATE_FORMAT = DateTimeFormatter.ISO_LOCAL_DATE;
    private static final Pattern FILE_DATE_PATTERN = Pattern.compile("(?<!\\d)(\\d{2})\\.(\\d{2})\\.(\\d{2}|\\d{4})(?!\\d)");
    private static final String SHEWHART_VLK_EVENT = "влк карта Шухарата по ЗИ";
    private static final String SHEWHART_VLK_MARK = "✓";

    private final DefaultListModel<File> inputFilesModel = new DefaultListModel<>();

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

        JLabel hint = new JLabel("Добавьте Excel-файлы кнопкой или перетащите в область ниже, затем нажмите «Преобразовать».", SwingConstants.LEFT);
        hint.setBorder(BorderFactory.createEmptyBorder(4, 0, 0, 0));

        panel.add(title, BorderLayout.NORTH);
        panel.add(hint, BorderLayout.SOUTH);
        panel.setBorder(BorderFactory.createEmptyBorder(8, 8, 0, 8));
        return panel;
    }

    private JComponent createFilesPanel() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));
        panel.setBorder(BorderFactory.createTitledBorder("Входные файлы"));

        JList<File> filesList = new JList<>(inputFilesModel);
        filesList.setCellRenderer((list, value, index, isSelected, cellHasFocus) -> {
            JLabel label = new JLabel(value.getAbsolutePath());
            if (isSelected) {
                label.setOpaque(true);
                label.setBackground(list.getSelectionBackground());
                label.setForeground(list.getSelectionForeground());
            }
            return label;
        });

        JScrollPane listScrollPane = new JScrollPane(filesList);
        listScrollPane.setPreferredSize(new Dimension(100, 220));

        JPanel dropPanel = new JPanel(new BorderLayout());
        dropPanel.setBorder(BorderFactory.createDashedBorder(Color.GRAY));
        dropPanel.setBackground(new Color(245, 245, 245));
        JLabel dropHint = new JLabel("Перетащите Excel-файлы сюда", SwingConstants.CENTER);
        dropPanel.add(dropHint, BorderLayout.CENTER);

        TransferHandler fileDropHandler = new TransferHandler() {
            @Override
            public boolean canImport(TransferSupport support) {
                return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            public boolean importData(TransferSupport support) {
                if (!canImport(support)) {
                    return false;
                }
                try {
                    @SuppressWarnings("unchecked")
                    List<File> droppedFiles = (List<File>) support.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    addExcelFiles(droppedFiles);
                    return true;
                } catch (Exception ignored) {
                    return false;
                }
            }
        };

        dropPanel.setTransferHandler(fileDropHandler);
        filesList.setTransferHandler(fileDropHandler);

        JButton addButton = new JButton("Добавить файлы");
        addButton.addActionListener(e -> chooseFiles());

        JButton removeButton = new JButton("Удалить выбранные");
        removeButton.addActionListener(e -> {
            List<File> selected = filesList.getSelectedValuesList();
            selected.forEach(inputFilesModel::removeElement);
        });

        JPanel controlsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        controlsPanel.add(addButton);
        controlsPanel.add(removeButton);

        panel.add(controlsPanel, BorderLayout.NORTH);
        panel.add(listScrollPane, BorderLayout.CENTER);
        panel.add(dropPanel, BorderLayout.SOUTH);
        return panel;
    }

    private JComponent createBottomPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 8));

        JButton clearButton = new JButton("Очистить");
        clearButton.addActionListener(e -> inputFilesModel.clear());

        JButton convertButton = new JButton("Преобразовать");
        convertButton.setFont(convertButton.getFont().deriveFont(Font.BOLD, 14f));
        convertButton.addActionListener(e -> convertFiles());

        panel.add(clearButton);
        panel.add(convertButton);
        return panel;
    }

    private void chooseFiles() {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Выберите Excel-файлы");
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xls", "xlsx", "xlsm"));

        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            addExcelFiles(List.of(chooser.getSelectedFiles()));
        }
    }

    private void addExcelFiles(List<File> files) {
        for (File file : files) {
            if (isExcelFile(file) && !containsFile(file)) {
                inputFilesModel.addElement(file);
            }
        }
    }

    private boolean containsFile(File file) {
        String absolutePath = file.getAbsolutePath();
        for (int i = 0; i < inputFilesModel.size(); i++) {
            if (inputFilesModel.get(i).getAbsolutePath().equalsIgnoreCase(absolutePath)) {
                return true;
            }
        }
        return false;
    }

    private boolean isExcelFile(File file) {
        String name = file.getName().toLowerCase();
        return name.endsWith(".xls") || name.endsWith(".xlsx") || name.endsWith(".xlsm");
    }

    private void convertFiles() {
        if (inputFilesModel.isEmpty()) {
            JOptionPane.showMessageDialog(
                    this,
                    "Добавьте хотя бы один Excel-файл для преобразования.",
                    "Нет файлов",
                    JOptionPane.WARNING_MESSAGE
            );
            return;
        }

        List<File> inputFiles = new ArrayList<>();
        for (int i = 0; i < inputFilesModel.size(); i++) {
            inputFiles.add(inputFilesModel.get(i));
        }

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Сохранить результат в формате XLSX");
        chooser.setSelectedFile(new File("Карта_Шухарта.xlsx"));
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Workbook", "xlsx"));

        int result = chooser.showSaveDialog(this);
        if (result != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File targetFile = chooser.getSelectedFile();
        if (!targetFile.getName().toLowerCase().endsWith(".xlsx")) {
            targetFile = new File(targetFile.getParentFile(), targetFile.getName() + ".xlsx");
        }

        try {
            ShewhartMapExcelExporter.exportStaticTitle(targetFile, inputFiles);

            int importedDates = saveShewhartDatesToVlk(inputFiles);
            String importedInfo = importedDates > 0
                    ? "\n\nДобавлено дат ВЛК: " + importedDates
                    : "\n\nДаты в названиях файлов не найдены или уже добавлены в БД.";

            JOptionPane.showMessageDialog(
                    this,
                    "Файл успешно создан:\n" + targetFile.getAbsolutePath() + importedInfo,
                    "Готово",
                    JOptionPane.INFORMATION_MESSAGE
            );
        } catch (IOException | SQLException ex) {
            JOptionPane.showMessageDialog(
                    this,
                    "Не удалось создать файл:\n" + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private int saveShewhartDatesToVlk(List<File> inputFiles) throws SQLException {
        Set<LocalDate> dates = new LinkedHashSet<>();
        for (File inputFile : inputFiles) {
            LocalDate parsed = extractDateFromFileName(inputFile.getName());
            if (parsed != null) {
                dates.add(parsed);
            }
        }

        if (dates.isEmpty()) {
            return 0;
        }

        List<VlkDateRecord> records = new ArrayList<>();
        for (LocalDate date : dates) {
            VlkDateRecord record = new VlkDateRecord();
            record.setVlkDate(date.format(DB_DATE_FORMAT));
            record.setResponsible(SHEWHART_VLK_MARK);
            record.setEventName(SHEWHART_VLK_EVENT);
            records.add(record);
        }

        return DatabaseManager.addVlkDatesIfMissing(records);
    }

    private LocalDate extractDateFromFileName(String fileName) {
        Matcher matcher = FILE_DATE_PATTERN.matcher(fileName);
        if (!matcher.find()) {
            return null;
        }

        int day = Integer.parseInt(matcher.group(1));
        int month = Integer.parseInt(matcher.group(2));
        String yearToken = matcher.group(3);
        int year = yearToken.length() == 2
                ? 2000 + Integer.parseInt(yearToken)
                : Integer.parseInt(yearToken);

        try {
            return LocalDate.of(year, month, day);
        } catch (RuntimeException ignored) {
            return null;
        }
    }
}
