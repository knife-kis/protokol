package ru.citlab24.protokol.tabs.qms;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.PersonnelRecord;
import ru.citlab24.protokol.db.VlkDateRecord;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.SQLException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Year;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;
import java.util.Set;

public class VlkTab extends JPanel {

    private static final Set<String> DEFAULT_PROTOCOL_EVENTS = Set.of(
            "Контроль точности результатов измерений по МИ Ш.13-2022",
            "Контроль точности результатов измерений по МР 2.6.1.0361-24",
            "Контроль точности результатов измерений по МУ 2.6.1.0333-23",
            "Контроль точности результатов измерений по «Методика измерения плотности потока радона с поверхности земли и строительных конструкций (свидетельство об аттестации МВИ № 40090.6К816)»",
            "Контроль точности результатов измерений по Руководство по эксплуатации Экофизика-110А ПКДУ.411000.001.02 РЭ п.7.1, п.7.2 с приложением МИ ПКФ 12-006 п.2, п.5",
            "Контроль точности результатов измерений по ГОСТ 27296-2012, ГОСТ Р ИСО 3382-2-2013, ГОСТ Р 56769-2015, ГОСТ Р 56770-2015",
            "Контроль точности результатов измерений по МИ СС.09-2021",
            "Контроль точности результатов измерений по ГОСТ 30494-2011"
    );

    private static final String[] EXTRA_TABLE_COLUMNS = {
            "Наименование мероприятий",
            "Периодичность",
            "Ответственный",
            "Формы учета"
    };

    private static final String[] VLK_DATES_COLUMNS = {
            "Дата ВЛК",
            "Ответственный",
            "Наименование мероприятия"
    };

    private static final DateTimeFormatter DB_DATE_FORMAT = DateTimeFormatter.ISO_LOCAL_DATE;
    private static final DateTimeFormatter UI_DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    private final CardLayout cardLayout = new CardLayout();
    private final JPanel cardPanel = new JPanel(cardLayout);

    private final JTextField yearField = new JTextField(String.valueOf(Year.now().getValue()), 6);
    private final JButton createPlanButton = new JButton("Сформировать Word-файл");
    private final JButton addRowButton = new JButton("+ Добавить строку");
    private final JButton swapResponsiblesButton = new JButton("Поменять Тарновский ↔ Белов");
    private final JButton vlkDatesButton = new JButton("Дата ВЛК");
    private final JLabel selectedYearLabel = new JLabel("Год плана не выбран", SwingConstants.LEFT);

    private final DefaultTableModel extraRowsModel = new DefaultTableModel(EXTRA_TABLE_COLUMNS, 0) {
        @Override
        public boolean isCellEditable(int row, int column) {
            return true;
        }
    };

    private final DefaultTableModel vlkDatesModel = new DefaultTableModel(VLK_DATES_COLUMNS, 0) {
        @Override
        public boolean isCellEditable(int row, int column) {
            return false;
        }
    };

    private final JTable vlkDatesTable = new JTable(vlkDatesModel);
    private final List<VlkDateRecord> vlkDateRecords = new ArrayList<>();
    private final Runnable onVlkDatesChanged;

    public VlkTab() {
        this(() -> {});
    }

    public VlkTab(Runnable onVlkDatesChanged) {
        super(new BorderLayout(10, 10));
        this.onVlkDatesChanged = onVlkDatesChanged == null ? () -> {} : onVlkDatesChanged;

        cardPanel.add(createMainView(), "main");
        cardPanel.add(createVlkDatesView(), "dates");
        add(cardPanel, BorderLayout.CENTER);

        fillDefaultExtraRows();
        wireActions();
        reloadVlkDates();
    }

    private JComponent createMainView() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.add(createToolbar(), BorderLayout.NORTH);
        panel.add(createExtraRowsPanel(), BorderLayout.CENTER);
        panel.add(selectedYearLabel, BorderLayout.SOUTH);
        return panel;
    }

    private JComponent createVlkDatesView() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));

        JPanel topBar = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        JButton backButton = new JButton("← Назад к ВЛК");
        JButton deleteButton = new JButton("Удалить запись");
        topBar.add(backButton);
        topBar.add(deleteButton);

        backButton.addActionListener(e -> cardLayout.show(cardPanel, "main"));
        deleteButton.addActionListener(e -> deleteSelectedVlkDate());

        vlkDatesTable.setRowHeight(28);
        vlkDatesTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        vlkDatesTable.getColumnModel().getColumn(0).setPreferredWidth(140);
        vlkDatesTable.getColumnModel().getColumn(1).setPreferredWidth(220);
        vlkDatesTable.getColumnModel().getColumn(2).setPreferredWidth(700);

        JScrollPane scrollPane = new JScrollPane(vlkDatesTable);
        scrollPane.setBorder(BorderFactory.createTitledBorder("Даты ВЛК"));

        panel.add(topBar, BorderLayout.NORTH);
        panel.add(scrollPane, BorderLayout.CENTER);
        return panel;
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

        gbc.gridx = 5;
        panel.add(vlkDatesButton, gbc);

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
        scrollPane.setBorder(BorderFactory.createTitledBorder("Дополнительные строки (редактируемые)"));
        scrollPane.getVerticalScrollBar().setUnitIncrement(24);
        return scrollPane;
    }

    private void wireActions() {
        addRowButton.addActionListener(e -> extraRowsModel.addRow(new String[]{"", "", "", ""}));
        swapResponsiblesButton.addActionListener(e -> swapResponsiblesInExtraRows());
        createPlanButton.addActionListener(e -> createPlanFile());
        vlkDatesButton.addActionListener(e -> {
            reloadVlkDates();
            cardLayout.show(cardPanel, "dates");
        });
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
                "Контроль точности результатов измерений по МР 2.6.1.0361-24",
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
        extraRowsModel.addRow(new String[]{
                "Контроль точности результатов измерений по ГОСТ 30494-2011",
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
        chooser.setDialogTitle("Выберите папку для документов ВЛК");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setSelectedFile(new File("ВЛК " + year));

        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }

        File baseDir = chooser.getSelectedFile();

        try {
            List<VlkWordExporter.PlanRow> extraRows = collectExtraRows();
            List<VlkWordExporter.PlanRow> allRows = new ArrayList<>(VlkWordExporter.mandatoryRows());
            allRows.addAll(extraRows);

            Path vlkDir = baseDir.toPath().resolve("ВЛК " + year);
            Files.createDirectories(vlkDir);

            File planFile = vlkDir.resolve("План_мониторинга_ВЛК_" + year + ".docx").toFile();
            VlkWordExporter.export(planFile, year, extraRows);

            int protocolCount = exportObservationProtocols(vlkDir, year, allRows);

            selectedYearLabel.setText("Сформирована папка: " + vlkDir);
            JOptionPane.showMessageDialog(this,
                    "Документы успешно созданы:\n" + vlkDir +
                            "\nПлан: " + planFile.getName() +
                            "\nПротоколов по результатам наблюдения: " + protocolCount);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(
                    this,
                    "Не удалось создать документ: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE
            );
        }
    }

    private int exportObservationProtocols(Path vlkDir,
                                           String year,
                                           List<VlkWordExporter.PlanRow> rows) throws IOException, SQLException {
        List<VlkWordExporter.PlanRow> protocolRows = rows.stream()
                .filter(row -> "Протокол по результатам наблюдения".equalsIgnoreCase(row.accountingForm().trim()))
                .filter(row -> DEFAULT_PROTOCOL_EVENTS.contains(row.event().trim()))
                .toList();

        List<LocalDate> protocolDates = generateProtocolDates(protocolRows.size(), Integer.parseInt(year));
        List<VlkDateRecord> generatedRecords = new ArrayList<>();
        Map<String, Integer> fileNameCounters = new HashMap<>();

        for (int i = 0; i < protocolRows.size(); i++) {
            VlkWordExporter.PlanRow row = protocolRows.get(i);
            LocalDate assignedDate = protocolDates.get(i);
            String displayDate = assignedDate.format(UI_DATE_FORMAT);

            String methodCipher = resolveMethodCipher(row.event());
            String safeMethodCipher = sanitizeFileNamePart(methodCipher);
            int fileIndex = fileNameCounters.merge(safeMethodCipher, 1, Integer::sum);
            String fileSuffix = fileIndex > 1 ? "_" + fileIndex : "";

            String protocolFileName = "Протокол_по_результатам_наблюдения_" + safeMethodCipher + fileSuffix + ".docx";
            ObservationProtocolWordExporter.export(vlkDir.resolve(protocolFileName).toFile(), year, row, displayDate);

            String issuanceFileName = "Лист_выдачи_приборов_" + safeMethodCipher + fileSuffix + ".docx";
            VlkEquipmentIssuanceWordExporter.export(vlkDir.resolve(issuanceFileName).toFile(), row);

            VlkDateRecord record = new VlkDateRecord();
            record.setVlkDate(assignedDate.format(DB_DATE_FORMAT));
            record.setResponsible(row.responsible());
            record.setEventName(row.event());
            generatedRecords.add(record);
        }

        DatabaseManager.replaceVlkDates(generatedRecords);
        reloadVlkDates();
        onVlkDatesChanged.run();
        return protocolRows.size();
    }

    private String resolveMethodCipher(String event) {
        String prefix = "Контроль точности результатов измерений по";
        if (event == null || event.isBlank()) {
            return "без_шифра";
        }
        if (event.startsWith(prefix)) {
            String cipher = event.substring(prefix.length()).trim();
            return cipher.isEmpty() ? "без_шифра" : cipher;
        }
        return event;
    }

    private String sanitizeFileNamePart(String value) {
        if (value == null || value.isBlank()) {
            return "без_шифра";
        }
        return value.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    private List<LocalDate> generateProtocolDates(int count, int year) throws SQLException {
        LocalDate from = LocalDate.of(year, 4, 15);
        LocalDate to = LocalDate.of(year, 8, 15);

        Set<LocalDate> unavailableDates = loadAllUnavailableDates();
        List<LocalDate> available = new ArrayList<>();

        for (LocalDate date = from; !date.isAfter(to); date = date.plusDays(1)) {
            if (date.getDayOfWeek() == DayOfWeek.SATURDAY || date.getDayOfWeek() == DayOfWeek.SUNDAY) {
                continue;
            }
            if (unavailableDates.contains(date)) {
                continue;
            }
            available.add(date);
        }

        if (available.isEmpty()) {
            throw new SQLException("Нет доступных дат ВЛК в периоде 15.04-15.08 (все даты заняты/выходные).");
        }

        Collections.shuffle(available, new Random());
        if (available.size() < count) {
            throw new SQLException("Недостаточно доступных дат ВЛК для всех протоколов: доступно " + available.size() + ", требуется " + count + ".");
        }

        return new ArrayList<>(available.subList(0, count));
    }

    private Set<LocalDate> loadAllUnavailableDates() throws SQLException {
        Set<LocalDate> result = new HashSet<>();
        for (PersonnelRecord person : DatabaseManager.getAllPersonnel()) {
            for (PersonnelRecord.UnavailabilityRecord rec : person.getUnavailabilityDates()) {
                String raw = rec.getUnavailableDate();
                if (raw == null || raw.isBlank()) {
                    continue;
                }
                try {
                    result.add(LocalDate.parse(raw.trim()));
                } catch (Exception ignored) {
                    // Пропускаем даты в некорректном формате
                }
            }
        }
        return result;
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

    private void reloadVlkDates() {
        vlkDateRecords.clear();
        try {
            vlkDateRecords.addAll(DatabaseManager.getAllVlkDates());
            vlkDatesModel.setRowCount(0);
            for (VlkDateRecord record : vlkDateRecords) {
                vlkDatesModel.addRow(new Object[]{
                        formatUiDate(record.getVlkDate()),
                        record.getResponsible(),
                        record.getEventName()
                });
            }
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось загрузить даты ВЛК: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void deleteSelectedVlkDate() {
        int selectedRow = vlkDatesTable.getSelectedRow();
        if (selectedRow < 0) {
            JOptionPane.showMessageDialog(this, "Выберите запись для удаления.");
            return;
        }

        VlkDateRecord record = vlkDateRecords.get(selectedRow);
        int confirm = JOptionPane.showConfirmDialog(
                this,
                "Точно удалить запись ВЛК за " + formatUiDate(record.getVlkDate()) + "?",
                "Подтверждение удаления",
                JOptionPane.YES_NO_OPTION
        );
        if (confirm != JOptionPane.YES_OPTION) {
            return;
        }

        try {
            DatabaseManager.deleteVlkDate(record.getId());
            reloadVlkDates();
            onVlkDatesChanged.run();
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось удалить запись: " + ex.getMessage(),
                    "Ошибка",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private String formatUiDate(String rawDate) {
        if (rawDate == null || rawDate.isBlank()) {
            return "";
        }
        try {
            return LocalDate.parse(rawDate).format(UI_DATE_FORMAT);
        } catch (Exception ex) {
            return rawDate;
        }
    }
}
