package ru.citlab24.protokol.tabs.resourceTab;

import com.github.lgooddatepicker.components.DatePicker;
import com.github.lgooddatepicker.components.DatePickerSettings;
import org.kordamp.ikonli.fontawesome5.FontAwesomeSolid;
import org.kordamp.ikonli.swing.FontIcon;
import ru.citlab24.protokol.db.AppUserRecord;
import ru.citlab24.protokol.equipment.EquipmentCategory;
import ru.citlab24.protokol.equipment.EquipmentContractRecord;
import ru.citlab24.protokol.equipment.EquipmentDetails;
import ru.citlab24.protokol.equipment.EquipmentEditorDialog;
import ru.citlab24.protokol.equipment.EquipmentFormExporter;
import ru.citlab24.protokol.equipment.EquipmentFormImporter;
import ru.citlab24.protokol.equipment.EquipmentHistoryDialog;
import ru.citlab24.protokol.equipment.EquipmentImportPreviewDialog;
import ru.citlab24.protokol.equipment.EquipmentRecord;
import ru.citlab24.protokol.equipment.EquipmentRecordFormatter;
import ru.citlab24.protokol.equipment.EquipmentRepository;
import ru.citlab24.protokol.equipment.EquipmentScheduleDialog;
import ru.citlab24.protokol.equipment.EquipmentServicePlanImportDialog;
import ru.citlab24.protokol.equipment.EquipmentServicePlanImporter;
import ru.citlab24.protokol.requests.RequestStorage;
import ru.citlab24.protokol.requests.StoredRequestFile;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableCellRenderer;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.IOException;
import java.net.URI;
import java.nio.file.Files;
import java.nio.file.Path;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import java.util.prefs.Preferences;

public class EquipmentTab extends JPanel {
    private static final String ALL_CATEGORIES = "Все виды";
    private static final String ALL_STATUSES = "Все статусы";
    private static final String ACTIVE_STATUS = "В эксплуатации";
    private static final String RETIRED_STATUS = "Выведено";
    private static final Preferences PREFERENCES =
            Preferences.userNodeForPackage(EquipmentTab.class);

    private final AppUserRecord currentUser;
    private final EquipmentTableModel tableModel = new EquipmentTableModel();
    private final JTable table = new JTable(tableModel);
    private final JTextField searchField = new JTextField(20);
    private final JComboBox<Object> categoryFilter = new JComboBox<>();
    private final JComboBox<String> statusFilter =
            new JComboBox<>(new String[]{ALL_STATUSES, ACTIVE_STATUS, RETIRED_STATUS});
    private final JCheckBox asOfCheck = new JCheckBox("Состояние на дату:");
    private final DatePicker asOfDatePicker = createDatePicker(LocalDate.now());
    private final JTextArea detailsArea = new JTextArea();
    private final JButton openControlLinkButton = new JButton("Открыть ФГИС",
            FontIcon.of(FontAwesomeSolid.FOLDER_OPEN, 14));
    private final EquipmentContractTableModel contractTableModel =
            new EquipmentContractTableModel();
    private final JTable contractTable = new JTable(contractTableModel);
    private final JLabel contractEquipmentLabel = new JLabel("Выберите оборудование");
    private final JButton addContractButton = new JButton("Добавить",
            FontIcon.of(FontAwesomeSolid.UPLOAD, 14));
    private final JButton openContractButton = new JButton("Открыть",
            FontIcon.of(FontAwesomeSolid.FOLDER_OPEN, 14));
    private final JButton replaceContractButton = new JButton("Заменить",
            FontIcon.of(FontAwesomeSolid.SYNC_ALT, 14));
    private final JButton deleteContractButton = new JButton("Удалить",
            FontIcon.of(FontAwesomeSolid.TRASH_ALT, 14));
    private final JButton openContractFolderButton = new JButton("Открыть папку",
            FontIcon.of(FontAwesomeSolid.FOLDER_OPEN, 14));
    private final EquipmentContractTableModel ownershipTableModel =
            new EquipmentContractTableModel();
    private final JTable ownershipTable = new JTable(ownershipTableModel);
    private final JLabel ownershipEquipmentLabel = new JLabel("Выберите оборудование");
    private final JButton addOwnershipButton = new JButton("Добавить",
            FontIcon.of(FontAwesomeSolid.UPLOAD, 14));
    private final JButton openOwnershipButton = new JButton("Открыть",
            FontIcon.of(FontAwesomeSolid.FOLDER_OPEN, 14));
    private final JButton replaceOwnershipButton = new JButton("Заменить",
            FontIcon.of(FontAwesomeSolid.SYNC_ALT, 14));
    private final JButton deleteOwnershipButton = new JButton("Удалить",
            FontIcon.of(FontAwesomeSolid.TRASH_ALT, 14));
    private final JButton openOwnershipFolderButton = new JButton("Открыть папку",
            FontIcon.of(FontAwesomeSolid.FOLDER_OPEN, 14));
    private final JLabel countLabel = new JLabel();
    private final List<EquipmentRecord> allRecords = new ArrayList<>();
    private RequestStorage requestStorage;

    public EquipmentTab(AppUserRecord currentUser) {
        super(new BorderLayout(8, 8));
        this.currentUser = currentUser;
        add(createHeader(), BorderLayout.NORTH);
        add(createContent(), BorderLayout.CENTER);
        add(countLabel, BorderLayout.SOUTH);
        configureTable();
        installListeners();
        reloadData();
    }

    private JComponent createHeader() {
        JPanel header = new JPanel(new BorderLayout(10, 8));

        JPanel actions = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        JLabel title = new JLabel("Реестр оборудования",
                FontIcon.of(FontAwesomeSolid.MICROSCOPE, 20), SwingConstants.LEFT);
        title.setFont(title.getFont().deriveFont(Font.BOLD, 16f));
        JButton formsButton = new JButton("Формы ▼",
                FontIcon.of(FontAwesomeSolid.FILE_WORD, 14));
        JPopupMenu formsMenu = new JPopupMenu();
        JMenuItem importFormsItem = new JMenuItem("Загрузить формы",
                FontIcon.of(FontAwesomeSolid.FILE_IMPORT, 14));
        JMenuItem exportFormsItem = new JMenuItem("Сформировать формы",
                FontIcon.of(FontAwesomeSolid.FILE_WORD, 14));
        JMenuItem importServicePlanItem = new JMenuItem("Загрузить текущий план Ф35",
                FontIcon.of(FontAwesomeSolid.FILE_IMPORT, 14));
        formsMenu.add(importFormsItem);
        formsMenu.add(exportFormsItem);
        formsMenu.addSeparator();
        formsMenu.add(importServicePlanItem);
        JButton scheduleButton = new JButton("Планы и графики Ф35–Ф38",
                FontIcon.of(FontAwesomeSolid.CALENDAR_ALT, 14));
        JButton addButton = new JButton("Добавить",
                FontIcon.of(FontAwesomeSolid.PLUS, 14));
        JButton editButton = new JButton("Редактировать",
                FontIcon.of(FontAwesomeSolid.EDIT, 14));
        JButton historyButton = new JButton("История",
                FontIcon.of(FontAwesomeSolid.HISTORY, 14));
        JButton deleteButton = new JButton("Удалить выбранное",
                FontIcon.of(FontAwesomeSolid.TRASH_ALT, 14));
        JButton refreshButton = new JButton("Обновить",
                FontIcon.of(FontAwesomeSolid.SYNC_ALT, 14));

        boolean canEdit = currentUser != null && currentUser.isAdministrator();
        importFormsItem.setEnabled(canEdit);
        importServicePlanItem.setEnabled(canEdit);
        addButton.setEnabled(canEdit);
        editButton.setEnabled(canEdit);
        deleteButton.setEnabled(canEdit);

        formsButton.addActionListener(event ->
                formsMenu.show(formsButton, 0, formsButton.getHeight()));
        importFormsItem.addActionListener(event -> importForms());
        exportFormsItem.addActionListener(event -> exportForms());
        importServicePlanItem.addActionListener(event -> importServicePlan());
        scheduleButton.addActionListener(event -> openEquipmentSchedules());
        addButton.addActionListener(event -> addEquipment());
        editButton.addActionListener(event -> editEquipment());
        historyButton.addActionListener(event -> showHistory());
        deleteButton.addActionListener(event -> deleteSelectedEquipment());
        refreshButton.addActionListener(event -> reloadData());

        actions.add(title);
        actions.add(Box.createHorizontalStrut(12));
        actions.add(formsButton);
        actions.add(scheduleButton);
        actions.add(addButton);
        actions.add(editButton);
        actions.add(historyButton);
        actions.add(deleteButton);
        actions.add(refreshButton);

        JPanel filters = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        categoryFilter.addItem(ALL_CATEGORIES);
        for (EquipmentCategory category : EquipmentCategory.values()) {
            categoryFilter.addItem(category);
        }
        filters.add(new JLabel("Поиск:"));
        filters.add(searchField);
        filters.add(new JLabel("Вид:"));
        filters.add(categoryFilter);
        filters.add(new JLabel("Статус:"));
        filters.add(statusFilter);
        filters.add(Box.createHorizontalStrut(10));
        filters.add(asOfCheck);
        filters.add(asOfDatePicker);

        header.add(actions, BorderLayout.NORTH);
        header.add(filters, BorderLayout.SOUTH);
        return header;
    }

    private JComponent createContent() {
        detailsArea.setEditable(false);
        detailsArea.setLineWrap(true);
        detailsArea.setWrapStyleWord(true);
        detailsArea.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));
        detailsArea.setFont(UIManager.getFont("TextArea.font"));

        JPanel detailsPanel = new JPanel(new BorderLayout(6, 6));
        JLabel detailsTitle = new JLabel("Карточка оборудования");
        detailsTitle.setFont(detailsTitle.getFont().deriveFont(Font.BOLD, 14f));
        JPanel detailsHeader = new JPanel(new BorderLayout(6, 0));
        detailsHeader.add(detailsTitle, BorderLayout.WEST);
        openControlLinkButton.setVisible(false);
        openControlLinkButton.addActionListener(event -> openControlLink());
        detailsHeader.add(openControlLinkButton, BorderLayout.EAST);
        detailsPanel.add(detailsHeader, BorderLayout.NORTH);
        detailsPanel.add(new JScrollPane(detailsArea), BorderLayout.CENTER);
        detailsPanel.setMinimumSize(new Dimension(320, 200));

        JTabbedPane sideTabs = new JTabbedPane();
        sideTabs.addTab("Карточка", detailsPanel);
        sideTabs.addTab("Договоры", createContractsPanel());
        sideTabs.addTab("Документы собственности", createOwnershipPanel());
        sideTabs.setMinimumSize(new Dimension(360, 200));

        JSplitPane split = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT,
                new JScrollPane(table), sideTabs);
        split.setResizeWeight(0.72);
        split.setDividerLocation(0.72);
        return split;
    }

    private JComponent createContractsPanel() {
        contractTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        contractTable.setAutoCreateRowSorter(true);
        contractTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        contractTable.setRowHeight(28);
        int[] widths = {230, 75, 115, 130};
        for (int index = 0; index < widths.length; index++) {
            contractTable.getColumnModel().getColumn(index).setPreferredWidth(widths[index]);
        }
        contractTable.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent event) {
                if (event.getClickCount() == 2) {
                    openSelectedContract();
                }
            }
        });

        addContractButton.addActionListener(event -> addContract());
        openContractButton.addActionListener(event -> openSelectedContract());
        replaceContractButton.addActionListener(event -> replaceSelectedContract());
        deleteContractButton.addActionListener(event -> deleteSelectedContract());
        openContractFolderButton.addActionListener(event -> openContractFolder());

        JPanel actions = new JPanel(new GridLayout(0, 2, 6, 6));
        actions.add(addContractButton);
        actions.add(openContractButton);
        actions.add(replaceContractButton);
        actions.add(deleteContractButton);
        actions.add(openContractFolderButton);
        actions.add(Box.createGlue());

        contractEquipmentLabel.setFont(contractEquipmentLabel.getFont().deriveFont(Font.BOLD));
        JPanel header = new JPanel(new BorderLayout(0, 8));
        header.setBorder(BorderFactory.createEmptyBorder(8, 8, 4, 8));
        header.add(contractEquipmentLabel, BorderLayout.NORTH);
        header.add(actions, BorderLayout.CENTER);

        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.add(header, BorderLayout.NORTH);
        panel.add(new JScrollPane(contractTable), BorderLayout.CENTER);
        updateContractActions();
        return panel;
    }

    private JComponent createOwnershipPanel() {
        ownershipTable.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        ownershipTable.setAutoCreateRowSorter(true);
        ownershipTable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        ownershipTable.setRowHeight(28);
        int[] widths = {230, 75, 115, 130};
        for (int index = 0; index < widths.length; index++) {
            ownershipTable.getColumnModel().getColumn(index).setPreferredWidth(widths[index]);
        }
        ownershipTable.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent event) {
                if (event.getClickCount() == 2) {
                    openSelectedOwnershipDocument();
                }
            }
        });

        addOwnershipButton.addActionListener(event -> addOwnershipDocument());
        openOwnershipButton.addActionListener(event -> openSelectedOwnershipDocument());
        replaceOwnershipButton.addActionListener(event -> replaceSelectedOwnershipDocument());
        deleteOwnershipButton.addActionListener(event -> deleteSelectedOwnershipDocument());
        openOwnershipFolderButton.addActionListener(event -> openOwnershipFolder());

        JPanel actions = new JPanel(new GridLayout(0, 2, 6, 6));
        actions.add(addOwnershipButton);
        actions.add(openOwnershipButton);
        actions.add(replaceOwnershipButton);
        actions.add(deleteOwnershipButton);
        actions.add(openOwnershipFolderButton);
        actions.add(Box.createGlue());

        ownershipEquipmentLabel.setFont(
                ownershipEquipmentLabel.getFont().deriveFont(Font.BOLD));
        JPanel header = new JPanel(new BorderLayout(0, 8));
        header.setBorder(BorderFactory.createEmptyBorder(8, 8, 4, 8));
        header.add(ownershipEquipmentLabel, BorderLayout.NORTH);
        header.add(actions, BorderLayout.CENTER);

        JPanel panel = new JPanel(new BorderLayout(0, 6));
        panel.add(header, BorderLayout.NORTH);
        panel.add(new JScrollPane(ownershipTable), BorderLayout.CENTER);
        updateOwnershipActions();
        return panel;
    }

    private void configureTable() {
        table.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        table.setAutoCreateRowSorter(true);
        table.setRowHeight(30);
        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        int[] widths = {50, 72, 300, 115, 145, 75, 185, 105, 110, 180, 135};
        for (int index = 0; index < widths.length; index++) {
            table.getColumnModel().getColumn(index).setPreferredWidth(widths[index]);
        }
        table.getColumnModel().getColumn(10).setCellRenderer(new StatusRenderer());
        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mouseClicked(MouseEvent event) {
                if (event.getClickCount() == 2
                        && currentUser != null && currentUser.isAdministrator()) {
                    editEquipment();
                }
            }
        });
    }

    private void installListeners() {
        table.getSelectionModel().addListSelectionListener(event -> {
            if (!event.getValueIsAdjusting()) {
                updateDetails(getSelectedRecord());
            }
        });
        contractTable.getSelectionModel().addListSelectionListener(event -> {
            if (!event.getValueIsAdjusting()) {
                updateContractActions();
            }
        });
        ownershipTable.getSelectionModel().addListSelectionListener(event -> {
            if (!event.getValueIsAdjusting()) {
                updateOwnershipActions();
            }
        });
        DocumentListener searchListener = new DocumentListener() {
            @Override
            public void insertUpdate(DocumentEvent event) {
                applyFilters();
            }

            @Override
            public void removeUpdate(DocumentEvent event) {
                applyFilters();
            }

            @Override
            public void changedUpdate(DocumentEvent event) {
                applyFilters();
            }
        };
        searchField.getDocument().addDocumentListener(searchListener);
        categoryFilter.addActionListener(event -> applyFilters());
        statusFilter.addActionListener(event -> applyFilters());
        asOfDatePicker.setEnabled(false);
        asOfCheck.addActionListener(event -> {
            asOfDatePicker.setEnabled(asOfCheck.isSelected());
            reloadData();
        });
        asOfDatePicker.addDateChangeListener(event -> {
            if (asOfCheck.isSelected() && event.getNewDate() != null) {
                reloadData();
            }
        });
    }

    public void reloadData() {
        Integer selectedId = getSelectedRecord() == null ? null : getSelectedRecord().equipmentId();
        try {
            LocalDate asOf = asOfCheck.isSelected() ? asOfDatePicker.getDate() : null;
            allRecords.clear();
            allRecords.addAll(EquipmentRepository.findAll(asOf));
            applyFilters();
            selectEquipment(selectedId);
        } catch (SQLException error) {
            showError("Не удалось загрузить реестр оборудования", error);
        }
    }

    private void applyFilters() {
        String query = normalize(searchField.getText());
        Object categoryValue = categoryFilter.getSelectedItem();
        String statusValue = Objects.toString(statusFilter.getSelectedItem(), ALL_STATUSES);
        List<EquipmentRecord> filtered = allRecords.stream()
                .filter(record -> !(categoryValue instanceof EquipmentCategory category)
                        || record.details().category() == category)
                .filter(record -> ALL_STATUSES.equals(statusValue)
                        || ACTIVE_STATUS.equals(statusValue) == record.details().active())
                .filter(record -> query.isBlank() || searchableText(record).contains(query))
                .toList();
        tableModel.setData(filtered, referenceDate());
        countLabel.setText("Оборудование: " + filtered.size()
                + (filtered.size() == allRecords.size() ? "" : " из " + allRecords.size()));
        if (!filtered.isEmpty()) {
            table.setRowSelectionInterval(0, 0);
        } else {
            updateDetails(null);
        }
    }

    private void importForms() {
        JFileChooser chooser = new ru.citlab24.protokol.ui.PathFileChooser(resolveImportDirectory());
        chooser.setDialogTitle("Выберите формы 2, 3 и 4");
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(new FileNameExtensionFilter("Документы Word (*.docx)", "docx"));
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File[] files = chooser.getSelectedFiles();
        if (files.length == 0 && chooser.getSelectedFile() != null) {
            files = new File[]{chooser.getSelectedFile()};
        }
        if (files.length == 0) {
            return;
        }
        File parent = files[0].getParentFile();
        if (parent != null) {
            PREFERENCES.put("equipment.import.directory", parent.getAbsolutePath());
        }

        try {
            List<EquipmentFormImporter.EquipmentImport> imports = new ArrayList<>();
            for (File file : files) {
                imports.add(EquipmentFormImporter.read(file.toPath()));
            }
            EquipmentImportPreviewDialog.ImportSelection selection =
                    EquipmentImportPreviewDialog.show(this, imports);
            if (selection == null) {
                return;
            }
            EquipmentRepository.ImportResult result = EquipmentRepository.importRows(
                    selection.rows(), selection.fullRegister(), currentUserId());
            reloadData();
            JOptionPane.showMessageDialog(this,
                    "Импорт завершён.\n"
                            + "Добавлено: " + result.added() + "\n"
                            + "Обновлено: " + result.updated() + "\n"
                            + "Выведено из эксплуатации: " + result.retired() + "\n"
                            + "Без изменений: " + result.unchanged(),
                    "Оборудование", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception error) {
            showError("Не удалось импортировать формы", error);
        }
    }

    private void exportForms() {
        LocalDate asOf = asOfCheck.isSelected() ? asOfDatePicker.getDate() : LocalDate.now();
        if (asOf == null) {
            JOptionPane.showMessageDialog(this, "Укажите дату состояния оборудования");
            return;
        }
        JFileChooser chooser = new ru.citlab24.protokol.ui.PathFileChooser(resolveExportDirectory());
        chooser.setDialogTitle("Выберите папку для форм 2, 3 и 4");
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setAcceptAllFileFilterUsed(false);
        chooser.setApproveButtonText("Сформировать");
        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION
                || chooser.getSelectedFile() == null) {
            return;
        }
        File directory = chooser.getSelectedFile();
        PREFERENCES.put("equipment.export.directory", directory.getAbsolutePath());
        try {
            List<EquipmentRecord> records = EquipmentRepository.findAll(asOf);
            List<Path> files = EquipmentFormExporter.export(
                    directory.toPath(), records, LocalDate.now());
            StringBuilder message = new StringBuilder(
                    "Формы сформированы по состоянию на "
                            + asOf.format(DateTimeFormatter.ofPattern("dd.MM.yyyy")) + ":\n");
            for (Path file : files) {
                message.append("\n").append(file.getFileName());
            }
            JOptionPane.showMessageDialog(this, message.toString(),
                    "Оборудование", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception error) {
            showError("Не удалось сформировать формы оборудования", error);
        }
    }

    private void importServicePlan() {
        JFileChooser chooser = new ru.citlab24.protokol.ui.PathFileChooser(resolveImportDirectory());
        chooser.setDialogTitle("Выберите текущий план обслуживания Ф35");
        chooser.setFileFilter(new FileNameExtensionFilter(
                "Документы Word (*.docx)", "docx"));
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION
                || chooser.getSelectedFile() == null) {
            return;
        }
        File file = chooser.getSelectedFile();
        File parent = file.getParentFile();
        if (parent != null) {
            PREFERENCES.put("equipment.import.directory", parent.getAbsolutePath());
        }

        try {
            EquipmentServicePlanImporter.ServicePlanImport plan =
                    EquipmentServicePlanImporter.read(file.toPath());
            List<EquipmentRecord> equipment = EquipmentRepository.findAll(null);
            EquipmentServicePlanImportDialog.ImportSelection selection =
                    EquipmentServicePlanImportDialog.show(this, plan, equipment);
            if (selection == null) {
                return;
            }

            Map<Integer, EquipmentRepository.ServicePlanUpdate> updates =
                    new LinkedHashMap<>();
            for (EquipmentServicePlanImportDialog.Assignment assignment
                    : selection.assignments()) {
                for (EquipmentRecord record : assignment.equipment()) {
                    EquipmentDetails details = record.details().toBuilder()
                            .effectiveDate(selection.effectiveDate())
                            .servicePlanRequired(true)
                            .servicePlanDetails(assignment.row().serviceDetails())
                            .servicePlanNotes(assignment.row().notes())
                            .build();
                    updates.put(record.equipmentId(),
                            new EquipmentRepository.ServicePlanUpdate(
                                    record.equipmentId(), details,
                                    assignment.row().sourceRow()));
                }
            }

            EquipmentRepository.ServicePlanImportResult result =
                    EquipmentRepository.importServicePlan(
                            new ArrayList<>(updates.values()),
                            file.getName(), currentUserId());
            reloadData();
            JOptionPane.showMessageDialog(this,
                    "План Ф35 загружен.\n"
                            + "Обновлено карточек: " + result.updated() + "\n"
                            + "Без изменений: " + result.unchanged() + "\n"
                            + "Строк без сопоставления: " + selection.skippedRows(),
                    "Оборудование", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception error) {
            showError("Не удалось загрузить текущий план Ф35", error);
        }
    }

    private void openEquipmentSchedules() {
        try {
            EquipmentScheduleDialog.show(this, currentUser, storage());
        } catch (IOException error) {
            showError("Не удалось открыть реестр графиков", error);
        }
    }

    private void addEquipment() {
        EquipmentDetails details = EquipmentEditorDialog.show(this,
                "Добавить оборудование", null);
        if (details == null) {
            return;
        }
        try {
            int equipmentId = EquipmentRepository.save(null, details, currentUserId());
            reloadData();
            selectEquipment(equipmentId);
            showCreationChecklist(details);
        } catch (SQLException error) {
            showError("Не удалось добавить оборудование", error);
        }
    }

    private void showCreationChecklist(EquipmentDetails details) {
        List<String> reminders = new ArrayList<>();
        if (details.servicePlanRequired()) {
            reminders.add("Сформировать новую редакцию плана обслуживания Ф35.");
        }
        if (details.maintenanceRequired()) {
            reminders.add("Сформировать новую редакцию графика технического обслуживания Ф36.");
        }
        if (details.category() == EquipmentCategory.MEASURING) {
            reminders.add("Сформировать новую редакцию графика поверки Ф37.");
        } else if (details.category() == EquipmentCategory.TESTING) {
            reminders.add("Сформировать новую редакцию графика аттестации Ф38.");
        }
        reminders.add("Внести данные во ФГИС в течение 10 рабочих дней.");

        StringBuilder message = new StringBuilder(
                "Оборудование сохранено. Не забудьте:\n\n");
        for (String reminder : reminders) {
            message.append("• ").append(reminder).append('\n');
        }
        Object[] options = {"Открыть планы и графики", "Закрыть"};
        int result = JOptionPane.showOptionDialog(
                this, message.toString(), "Следующие действия",
                JOptionPane.DEFAULT_OPTION, JOptionPane.INFORMATION_MESSAGE,
                null, options, options[0]);
        if (result == 0) {
            openEquipmentSchedules();
        }
    }

    private void editEquipment() {
        EquipmentRecord selected = getSelectedRecord();
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        EquipmentDetails details = EquipmentEditorDialog.show(this,
                "Редактировать оборудование", selected.details());
        if (details == null) {
            return;
        }
        try {
            EquipmentRepository.save(selected.equipmentId(), details, currentUserId());
            reloadData();
            selectEquipment(selected.equipmentId());
        } catch (SQLException error) {
            showError("Не удалось сохранить изменения", error);
        }
    }

    private void deleteSelectedEquipment() {
        List<EquipmentRecord> selected = getSelectedRecords();
        if (selected.isEmpty()) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование для удаления");
            return;
        }

        List<EquipmentContractRecord> documents = new ArrayList<>();
        try {
            for (EquipmentRecord record : selected) {
                documents.addAll(EquipmentRepository.findContracts(record.equipmentId()));
                documents.addAll(EquipmentRepository.findOwnershipDocuments(record.equipmentId()));
            }
        } catch (SQLException error) {
            showError("Не удалось проверить связанные документы", error);
            return;
        }

        String target = selected.size() == 1
                ? "оборудование «" + selected.get(0).details().nameType() + "»"
                : "выбранное оборудование (" + selected.size() + " записей)";
        String message = "Вы уверены, что хотите удалить " + target + "?\n\n"
                + "Будут удалены история изменений"
                + (documents.isEmpty()
                ? "."
                : " и связанные файлы (" + documents.size() + ").")
                + "\nЭто действие нельзя отменить.";
        Object[] options = {"Удалить", "Отмена"};
        int answer = JOptionPane.showOptionDialog(
                this, message, "Удаление оборудования",
                JOptionPane.DEFAULT_OPTION, JOptionPane.WARNING_MESSAGE,
                null, options, options[1]);
        if (answer != 0) {
            return;
        }

        try {
            int deleted = EquipmentRepository.deleteEquipment(
                    selected.stream().map(EquipmentRecord::equipmentId).toList());
            List<String> fileErrors = new ArrayList<>();
            if (!documents.isEmpty()) {
                try {
                    RequestStorage files = storage();
                    for (EquipmentContractRecord document : documents) {
                        try {
                            files.delete(document.relativePath());
                        } catch (IOException error) {
                            fileErrors.add(document.originalFileName());
                        }
                    }
                } catch (IOException error) {
                    for (EquipmentContractRecord document : documents) {
                        fileErrors.add(document.originalFileName());
                    }
                }
            }
            reloadData();
            if (fileErrors.isEmpty()) {
                JOptionPane.showMessageDialog(this,
                        "Удалено оборудования: " + deleted,
                        "Оборудование", JOptionPane.INFORMATION_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(this,
                        "Удалено оборудования: " + deleted
                                + "\nНе удалось удалить связанных файлов: "
                                + fileErrors.size(),
                        "Оборудование", JOptionPane.WARNING_MESSAGE);
            }
        } catch (SQLException error) {
            showError("Не удалось удалить оборудование", error);
        }
    }

    private void showHistory() {
        EquipmentRecord selected = getSelectedRecord();
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        try {
            EquipmentHistoryDialog.show(this, selected,
                    EquipmentRepository.findHistory(selected.equipmentId()));
        } catch (SQLException error) {
            showError("Не удалось загрузить историю оборудования", error);
        }
    }

    private EquipmentRecord getSelectedRecord() {
        int viewRow = table.getSelectedRow();
        if (viewRow < 0) {
            return null;
        }
        return tableModel.getAt(table.convertRowIndexToModel(viewRow));
    }

    private List<EquipmentRecord> getSelectedRecords() {
        int[] viewRows = table.getSelectedRows();
        List<EquipmentRecord> selected = new ArrayList<>(viewRows.length);
        for (int viewRow : viewRows) {
            selected.add(tableModel.getAt(table.convertRowIndexToModel(viewRow)));
        }
        return selected;
    }

    private void selectEquipment(Integer equipmentId) {
        if (equipmentId == null) {
            return;
        }
        int modelRow = tableModel.indexOf(equipmentId);
        if (modelRow >= 0) {
            int viewRow = table.convertRowIndexToView(modelRow);
            table.setRowSelectionInterval(viewRow, viewRow);
            table.scrollRectToVisible(table.getCellRect(viewRow, 0, true));
        }
    }

    private void updateDetails(EquipmentRecord record) {
        detailsArea.setText(EquipmentRecordFormatter.fullText(record));
        detailsArea.setCaretPosition(0);
        boolean hasLink = record != null && !record.details().controlLink().isBlank();
        openControlLinkButton.setVisible(hasLink);
        reloadContracts(record);
        reloadOwnershipDocuments(record);
    }

    private void reloadContracts(EquipmentRecord record) {
        if (record == null) {
            contractEquipmentLabel.setText("Выберите оборудование");
            contractTableModel.setData(List.of());
            updateContractActions();
            return;
        }
        contractEquipmentLabel.setText(record.details().nameType());
        try {
            contractTableModel.setData(
                    EquipmentRepository.findContracts(record.equipmentId()));
        } catch (SQLException error) {
            contractTableModel.setData(List.of());
            showError("Не удалось загрузить договоры оборудования", error);
        }
        updateContractActions();
    }

    private void addContract() {
        EquipmentRecord equipment = getSelectedRecord();
        if (equipment == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        File source = chooseContractFile("Добавить договор оборудования");
        if (source == null) {
            return;
        }
        StoredRequestFile stored = null;
        try {
            RequestStorage storage = storage();
            stored = storage.storeEquipmentContract(
                    source.toPath(), equipment.equipmentId(), null, null);
            EquipmentRepository.addContract(
                    equipment.equipmentId(), stored, currentUserId());
            reloadContracts(equipment);
        } catch (Exception error) {
            if (stored != null) {
                try {
                    storage().delete(stored.relativePath());
                } catch (IOException ignored) {
                }
            }
            showError("Не удалось добавить договор оборудования", error);
        }
    }

    private void openSelectedContract() {
        EquipmentContractRecord contract = getSelectedContract();
        if (contract == null) {
            JOptionPane.showMessageDialog(this, "Выберите договор");
            return;
        }
        try {
            storage().open(contract.relativePath());
        } catch (Exception error) {
            showError("Не удалось открыть договор оборудования", error);
        }
    }

    private void replaceSelectedContract() {
        EquipmentRecord equipment = getSelectedRecord();
        EquipmentContractRecord contract = getSelectedContract();
        if (equipment == null || contract == null) {
            JOptionPane.showMessageDialog(this, "Выберите договор");
            return;
        }
        File source = chooseContractFile("Заменить договор оборудования");
        if (source == null) {
            return;
        }
        StoredRequestFile stored = null;
        boolean databaseUpdated = false;
        try {
            RequestStorage storage = storage();
            stored = storage.storeEquipmentContract(
                    source.toPath(), equipment.equipmentId(), null, null);
            EquipmentRepository.replaceContract(contract.id(), stored, currentUserId());
            databaseUpdated = true;
            try {
                storage.delete(contract.relativePath());
            } catch (IOException error) {
                JOptionPane.showMessageDialog(this,
                        "Договор заменён, но старый файл не удалось удалить:\n"
                                + error.getMessage(),
                        "Договоры оборудования", JOptionPane.WARNING_MESSAGE);
            }
            reloadContracts(equipment);
        } catch (Exception error) {
            if (!databaseUpdated && stored != null) {
                try {
                    storage().delete(stored.relativePath());
                } catch (IOException ignored) {
                }
            }
            showError("Не удалось заменить договор оборудования", error);
        }
    }

    private void deleteSelectedContract() {
        EquipmentRecord equipment = getSelectedRecord();
        EquipmentContractRecord contract = getSelectedContract();
        if (equipment == null || contract == null) {
            JOptionPane.showMessageDialog(this, "Выберите договор");
            return;
        }
        int answer = JOptionPane.showConfirmDialog(this,
                "Удалить договор «" + contract.originalFileName() + "»?",
                "Договоры оборудования",
                JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer != JOptionPane.YES_OPTION) {
            return;
        }
        try {
            RequestStorage storage = storage();
            EquipmentContractRecord removed =
                    EquipmentRepository.deleteContract(contract.id());
            try {
                storage.delete(removed.relativePath());
            } catch (IOException error) {
                JOptionPane.showMessageDialog(this,
                        "Запись удалена, но файл не удалось удалить:\n"
                                + error.getMessage(),
                        "Договоры оборудования", JOptionPane.WARNING_MESSAGE);
            }
            reloadContracts(equipment);
        } catch (Exception error) {
            showError("Не удалось удалить договор оборудования", error);
        }
    }

    private void openContractFolder() {
        EquipmentRecord equipment = getSelectedRecord();
        if (equipment == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        try {
            storage().openEquipmentContractFolder(equipment.equipmentId());
        } catch (Exception error) {
            showError("Не удалось открыть папку договоров", error);
        }
    }

    private EquipmentContractRecord getSelectedContract() {
        int viewRow = contractTable.getSelectedRow();
        if (viewRow < 0) {
            return null;
        }
        return contractTableModel.getAt(contractTable.convertRowIndexToModel(viewRow));
    }

    private void updateContractActions() {
        boolean equipmentSelected = getSelectedRecord() != null;
        boolean contractSelected = getSelectedContract() != null;
        boolean canEdit = currentUser != null && currentUser.isAdministrator();
        addContractButton.setEnabled(canEdit && equipmentSelected);
        openContractButton.setEnabled(contractSelected);
        replaceContractButton.setEnabled(canEdit && contractSelected);
        deleteContractButton.setEnabled(canEdit && contractSelected);
        openContractFolderButton.setEnabled(equipmentSelected);
    }

    private File chooseContractFile(String title) {
        JFileChooser chooser = new ru.citlab24.protokol.ui.PathFileChooser(resolveContractDirectory());
        chooser.setDialogTitle(title);
        chooser.setFileFilter(new FileNameExtensionFilter(
                "Документы (*.doc, *.docx, *.pdf, *.xls, *.xlsx, *.rtf)",
                "doc", "docx", "pdf", "xls", "xlsx", "rtf"));
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION
                || chooser.getSelectedFile() == null) {
            return null;
        }
        File selected = chooser.getSelectedFile();
        File parent = selected.getParentFile();
        if (parent != null) {
            PREFERENCES.put("equipment.contract.directory", parent.getAbsolutePath());
        }
        return selected;
    }

    private void reloadOwnershipDocuments(EquipmentRecord record) {
        if (record == null) {
            ownershipEquipmentLabel.setText("Выберите оборудование");
            ownershipTableModel.setData(List.of());
            updateOwnershipActions();
            return;
        }
        ownershipEquipmentLabel.setText(record.details().nameType());
        try {
            ownershipTableModel.setData(
                    EquipmentRepository.findOwnershipDocuments(record.equipmentId()));
        } catch (SQLException error) {
            ownershipTableModel.setData(List.of());
            showError("Не удалось загрузить документы собственности", error);
        }
        updateOwnershipActions();
    }

    private void addOwnershipDocument() {
        EquipmentRecord equipment = getSelectedRecord();
        if (equipment == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        File source = chooseOwnershipFile("Добавить документ собственности");
        if (source == null) {
            return;
        }
        StoredRequestFile stored = null;
        try {
            RequestStorage storage = storage();
            stored = storage.storeEquipmentOwnershipDocument(
                    source.toPath(), equipment.equipmentId(), null, null);
            EquipmentRepository.addOwnershipDocument(
                    equipment.equipmentId(), stored, currentUserId());
            reloadOwnershipDocuments(equipment);
        } catch (Exception error) {
            if (stored != null) {
                try {
                    storage().delete(stored.relativePath());
                } catch (IOException ignored) {
                }
            }
            showError("Не удалось добавить документ собственности", error);
        }
    }

    private void openSelectedOwnershipDocument() {
        EquipmentContractRecord document = getSelectedOwnershipDocument();
        if (document == null) {
            JOptionPane.showMessageDialog(this, "Выберите документ");
            return;
        }
        try {
            storage().open(document.relativePath());
        } catch (Exception error) {
            showError("Не удалось открыть документ собственности", error);
        }
    }

    private void replaceSelectedOwnershipDocument() {
        EquipmentRecord equipment = getSelectedRecord();
        EquipmentContractRecord document = getSelectedOwnershipDocument();
        if (equipment == null || document == null) {
            JOptionPane.showMessageDialog(this, "Выберите документ");
            return;
        }
        File source = chooseOwnershipFile("Заменить документ собственности");
        if (source == null) {
            return;
        }
        StoredRequestFile stored = null;
        boolean databaseUpdated = false;
        try {
            RequestStorage storage = storage();
            stored = storage.storeEquipmentOwnershipDocument(
                    source.toPath(), equipment.equipmentId(), null, null);
            EquipmentRepository.replaceContract(document.id(), stored, currentUserId());
            databaseUpdated = true;
            try {
                storage.delete(document.relativePath());
            } catch (IOException error) {
                JOptionPane.showMessageDialog(this,
                        "Документ заменён, но старый файл не удалось удалить:\n"
                                + error.getMessage(),
                        "Документы собственности", JOptionPane.WARNING_MESSAGE);
            }
            reloadOwnershipDocuments(equipment);
        } catch (Exception error) {
            if (!databaseUpdated && stored != null) {
                try {
                    storage().delete(stored.relativePath());
                } catch (IOException ignored) {
                }
            }
            showError("Не удалось заменить документ собственности", error);
        }
    }

    private void deleteSelectedOwnershipDocument() {
        EquipmentRecord equipment = getSelectedRecord();
        EquipmentContractRecord document = getSelectedOwnershipDocument();
        if (equipment == null || document == null) {
            JOptionPane.showMessageDialog(this, "Выберите документ");
            return;
        }
        int answer = JOptionPane.showConfirmDialog(this,
                "Удалить документ «" + document.originalFileName() + "»?",
                "Документы собственности",
                JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
        if (answer != JOptionPane.YES_OPTION) {
            return;
        }
        try {
            RequestStorage storage = storage();
            EquipmentContractRecord removed =
                    EquipmentRepository.deleteContract(document.id());
            try {
                storage.delete(removed.relativePath());
            } catch (IOException error) {
                JOptionPane.showMessageDialog(this,
                        "Запись удалена, но файл не удалось удалить:\n"
                                + error.getMessage(),
                        "Документы собственности", JOptionPane.WARNING_MESSAGE);
            }
            reloadOwnershipDocuments(equipment);
        } catch (Exception error) {
            showError("Не удалось удалить документ собственности", error);
        }
    }

    private void openOwnershipFolder() {
        EquipmentRecord equipment = getSelectedRecord();
        if (equipment == null) {
            JOptionPane.showMessageDialog(this, "Выберите оборудование");
            return;
        }
        try {
            storage().openEquipmentOwnershipFolder(equipment.equipmentId());
        } catch (Exception error) {
            showError("Не удалось открыть папку документов собственности", error);
        }
    }

    private EquipmentContractRecord getSelectedOwnershipDocument() {
        int viewRow = ownershipTable.getSelectedRow();
        if (viewRow < 0) {
            return null;
        }
        return ownershipTableModel.getAt(
                ownershipTable.convertRowIndexToModel(viewRow));
    }

    private void updateOwnershipActions() {
        boolean equipmentSelected = getSelectedRecord() != null;
        boolean documentSelected = getSelectedOwnershipDocument() != null;
        boolean canEdit = currentUser != null && currentUser.isAdministrator();
        addOwnershipButton.setEnabled(canEdit && equipmentSelected);
        openOwnershipButton.setEnabled(documentSelected);
        replaceOwnershipButton.setEnabled(canEdit && documentSelected);
        deleteOwnershipButton.setEnabled(canEdit && documentSelected);
        openOwnershipFolderButton.setEnabled(equipmentSelected);
    }

    private File chooseOwnershipFile(String title) {
        JFileChooser chooser = new ru.citlab24.protokol.ui.PathFileChooser(resolveOwnershipDirectory());
        chooser.setDialogTitle(title);
        chooser.setFileFilter(new FileNameExtensionFilter(
                "Документы и сканы (*.pdf, *.jpg, *.jpeg, *.png, *.tif, *.tiff, *.doc, *.docx, *.xls, *.xlsx)",
                "pdf", "jpg", "jpeg", "png", "tif", "tiff",
                "doc", "docx", "xls", "xlsx"));
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION
                || chooser.getSelectedFile() == null) {
            return null;
        }
        File selected = chooser.getSelectedFile();
        File parent = selected.getParentFile();
        if (parent != null) {
            PREFERENCES.put("equipment.ownership.directory", parent.getAbsolutePath());
        }
        return selected;
    }

    private File resolveOwnershipDirectory() {
        String saved = PREFERENCES.get("equipment.ownership.directory", "");
        if (!saved.isBlank() && Files.isDirectory(Path.of(saved))) {
            return new File(saved);
        }
        return resolveContractDirectory();
    }

    private File resolveContractDirectory() {
        String saved = PREFERENCES.get("equipment.contract.directory", "");
        if (!saved.isBlank() && Files.isDirectory(Path.of(saved))) {
            return new File(saved);
        }
        return resolveImportDirectory();
    }

    private RequestStorage storage() throws IOException {
        if (requestStorage == null) {
            requestStorage = new RequestStorage();
        }
        return requestStorage;
    }

    private void openControlLink() {
        EquipmentRecord selected = getSelectedRecord();
        if (selected == null || selected.details().controlLink().isBlank()) {
            return;
        }
        try {
            Desktop.getDesktop().browse(URI.create(selected.details().controlLink()));
        } catch (Exception error) {
            showError("Не удалось открыть ссылку ФГИС", error);
        }
    }

    private File resolveImportDirectory() {
        String saved = PREFERENCES.get("equipment.import.directory", "");
        if (!saved.isBlank() && Files.isDirectory(Path.of(saved))) {
            return new File(saved);
        }
        return new File(System.getProperty("user.home"), "Documents");
    }

    private File resolveExportDirectory() {
        String saved = PREFERENCES.get("equipment.export.directory", "");
        if (!saved.isBlank() && Files.isDirectory(Path.of(saved))) {
            return new File(saved);
        }
        return resolveImportDirectory();
    }

    private Integer currentUserId() {
        return currentUser == null ? null : currentUser.getId();
    }

    private LocalDate referenceDate() {
        LocalDate selected = asOfCheck.isSelected() ? asOfDatePicker.getDate() : null;
        return selected == null ? LocalDate.now() : selected;
    }

    private static String searchableText(EquipmentRecord record) {
        EquipmentDetails value = record.details();
        return normalize(String.join(" ",
                value.category().getTitle(), value.measuredCharacteristics(),
                value.testedObjectGroups(), value.nameType(), value.registryNumber(),
                value.completeness(), value.manufacturer(),
                Objects.toString(value.commissioningYear(), ""),
                value.factoryNumber(), Objects.toString(value.inventoryNumber(), ""),
                value.identification(), value.controlNumber(), value.verificationPlace(),
                value.ownershipDocument(), value.storageLocation(),
                value.servicePlanDetails(), value.servicePlanNotes(), value.notes()));
    }

    private static String normalize(String value) {
        return value == null ? "" : value.toLowerCase(Locale.ROOT)
                .replace('ё', 'е')
                .replaceAll("\\s+", " ")
                .trim();
    }

    private void showError(String title, Exception error) {
        JOptionPane.showMessageDialog(this,
                title + ":\n" + error.getMessage(),
                "Ошибка", JOptionPane.ERROR_MESSAGE);
    }

    private static DatePicker createDatePicker(LocalDate initialDate) {
        DatePickerSettings settings = new DatePickerSettings(new Locale("ru", "RU"));
        settings.setFormatForDatesCommonEra("dd.MM.yyyy");
        DatePicker picker = new DatePicker(settings);
        picker.setDate(initialDate);
        picker.setPreferredSize(new Dimension(155, 30));
        return picker;
    }

    private static String equipmentStatus(EquipmentDetails value, LocalDate referenceDate) {
        if (!value.active()) {
            return RETIRED_STATUS;
        }
        if (value.category() == EquipmentCategory.AUXILIARY
                || value.controlValidUntil() == null) {
            return ACTIVE_STATUS;
        }
        if (value.controlValidUntil().isBefore(referenceDate)) {
            return "Срок истёк";
        }
        if (!value.controlValidUntil().isAfter(referenceDate.plusDays(60))) {
            return "Срок скоро истечёт";
        }
        return "Действует";
    }

    private static final class EquipmentTableModel extends AbstractTableModel {
        private static final DateTimeFormatter DATE_FORMAT =
                DateTimeFormatter.ofPattern("dd.MM.yyyy");
        private final String[] columns = {
                "№", "Вид", "Наименование и тип", "№ Госреестра",
                "Зав. №", "Инв. №", "Поверка / аттестация", "Дата",
                "Действует до", "Место хранения", "Статус"
        };
        private List<EquipmentRecord> data = List.of();
        private LocalDate referenceDate = LocalDate.now();

        void setData(List<EquipmentRecord> rows, LocalDate date) {
            data = new ArrayList<>(rows);
            referenceDate = date;
            fireTableDataChanged();
        }

        EquipmentRecord getAt(int row) {
            return data.get(row);
        }

        int indexOf(int equipmentId) {
            for (int index = 0; index < data.size(); index++) {
                if (data.get(index).equipmentId() == equipmentId) {
                    return index;
                }
            }
            return -1;
        }

        @Override
        public int getRowCount() {
            return data.size();
        }

        @Override
        public int getColumnCount() {
            return columns.length;
        }

        @Override
        public String getColumnName(int column) {
            return columns[column];
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            EquipmentDetails value = data.get(rowIndex).details();
            return switch (columnIndex) {
                case 0 -> value.position() > 0 ? value.position() : rowIndex + 1;
                case 1 -> value.category().getShortTitle();
                case 2 -> oneLine(value.nameType());
                case 3 -> value.registryNumber();
                case 4 -> oneLine(value.factoryNumber());
                case 5 -> value.inventoryNumber() == null
                        ? "" : value.inventoryNumber();
                case 6 -> value.controlNumber();
                case 7 -> formatDate(value.controlDate());
                case 8 -> formatDate(value.controlValidUntil());
                case 9 -> oneLine(value.storageLocation());
                case 10 -> equipmentStatus(value, referenceDate);
                default -> "";
            };
        }

        private static String oneLine(String value) {
            return value == null ? "" : value.replace('\n', ' ').replaceAll("\\s+", " ").trim();
        }

        private static String formatDate(LocalDate value) {
            return value == null ? "" : DATE_FORMAT.format(value);
        }
    }

    private static final class EquipmentContractTableModel extends AbstractTableModel {
        private static final DateTimeFormatter DATE_TIME_FORMAT =
                DateTimeFormatter.ofPattern("dd.MM.yyyy HH:mm");
        private final String[] columns = {"Файл", "Размер", "Добавил", "Добавлен"};
        private List<EquipmentContractRecord> data = List.of();

        void setData(List<EquipmentContractRecord> contracts) {
            data = new ArrayList<>(contracts);
            fireTableDataChanged();
        }

        EquipmentContractRecord getAt(int row) {
            return data.get(row);
        }

        @Override
        public int getRowCount() {
            return data.size();
        }

        @Override
        public int getColumnCount() {
            return columns.length;
        }

        @Override
        public String getColumnName(int column) {
            return columns[column];
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            EquipmentContractRecord contract = data.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> contract.originalFileName();
                case 1 -> formatSize(contract.sizeBytes());
                case 2 -> contract.uploadedBy();
                case 3 -> contract.uploadedAt().equals(java.time.LocalDateTime.MIN)
                        ? "" : contract.uploadedAt().format(DATE_TIME_FORMAT);
                default -> "";
            };
        }

        private static String formatSize(long size) {
            if (size < 1024) {
                return size + " Б";
            }
            if (size < 1024L * 1024L) {
                return String.format(Locale.ROOT, "%.1f КБ", size / 1024.0);
            }
            return String.format(Locale.ROOT, "%.1f МБ", size / (1024.0 * 1024.0));
        }
    }

    private static final class StatusRenderer extends DefaultTableCellRenderer {
        private static final Color GREEN = new Color(0xDDF2E2);
        private static final Color YELLOW = new Color(0xFFF2C2);
        private static final Color RED = new Color(0xF8D7DA);
        private static final Color GRAY = new Color(0xE9ECEF);

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean selected, boolean focus,
                                                       int row, int column) {
            Component component = super.getTableCellRendererComponent(
                    table, value, selected, focus, row, column);
            if (!selected) {
                String status = Objects.toString(value, "");
                component.setForeground(Color.DARK_GRAY);
                component.setBackground(switch (status) {
                    case "Действует", ACTIVE_STATUS -> GREEN;
                    case "Срок скоро истечёт" -> YELLOW;
                    case "Срок истёк" -> RED;
                    default -> GRAY;
                });
            }
            setHorizontalAlignment(SwingConstants.CENTER);
            return component;
        }
    }
}
