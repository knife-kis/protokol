package ru.citlab24.protokol.tabs.resourceTab;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.PersonnelRecord;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import java.awt.*;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

public class PersonnelTab extends JPanel {
    private static final DateTimeFormatter UI_DATE_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");
    private static final String DATE_PLACEHOLDER = "дд-мм-гггг";
    private static final String YEAR_FILTER_ALL = "Все годы";

    private final PersonnelTableModel personnelModel = new PersonnelTableModel();
    private final JTable personnelTable = new JTable(personnelModel);

    private final UnavailabilityTableModel unavailabilityModel = new UnavailabilityTableModel();
    private final JTable unavailabilityTable = new JTable(unavailabilityModel);
    private final JComboBox<String> yearFilterCombo = new JComboBox<>();

    private final JTextField searchField = new JTextField();
    private final List<PersonnelRecord> allPersonnel = new ArrayList<>();

    public PersonnelTab() {
        super(new BorderLayout(8, 8));
        add(createToolbar(), BorderLayout.NORTH);
        add(createContent(), BorderLayout.CENTER);

        personnelTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                refreshUnavailabilityPanel(getSelectedPerson());
            }
        });

        yearFilterCombo.addActionListener(e -> {
            PersonnelRecord selected = getSelectedPerson();
            if (selected != null) {
                applyUnavailabilityFilter(selected);
            }
        });

        reloadPersonnel();
    }

    private JComponent createToolbar() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));

        JPanel buttons = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        JButton addPersonBtn = new JButton("Добавить человека");
        JButton editPersonBtn = new JButton("Редактировать");
        JButton deletePersonBtn = new JButton("Удалить");
        JButton addDateBtn = new JButton("Добавить дату занятости");
        JButton deleteDateBtn = new JButton("Удалить дату");

        addPersonBtn.addActionListener(e -> addPerson());
        editPersonBtn.addActionListener(e -> editPerson());
        deletePersonBtn.addActionListener(e -> deletePerson());
        addDateBtn.addActionListener(e -> addUnavailableDate());
        deleteDateBtn.addActionListener(e -> deleteUnavailableDate());

        buttons.add(addPersonBtn);
        buttons.add(editPersonBtn);
        buttons.add(deletePersonBtn);
        buttons.add(addDateBtn);
        buttons.add(deleteDateBtn);

        JPanel searchPanel = new JPanel(new BorderLayout(4, 0));
        searchPanel.add(new JLabel("Поиск:"), BorderLayout.WEST);
        searchPanel.add(searchField, BorderLayout.CENTER);
        JButton findButton = new JButton("Найти");
        findButton.addActionListener(e -> applyFilter());
        searchField.addActionListener(e -> applyFilter());
        searchPanel.add(findButton, BorderLayout.EAST);

        panel.add(buttons, BorderLayout.WEST);
        panel.add(searchPanel, BorderLayout.EAST);
        return panel;
    }

    private JComponent createContent() {
        JPanel bottomPanel = new JPanel(new BorderLayout(6, 6));
        bottomPanel.add(new JScrollPane(unavailabilityTable), BorderLayout.CENTER);

        JPanel filterPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        filterPanel.add(new JLabel("Год:"));
        filterPanel.add(yearFilterCombo);
        bottomPanel.add(filterPanel, BorderLayout.SOUTH);

        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
                new JScrollPane(personnelTable),
                bottomPanel);
        split.setResizeWeight(0.6);
        return split;
    }

    private void reloadPersonnel() {
        try {
            allPersonnel.clear();
            allPersonnel.addAll(DatabaseManager.getAllPersonnel());
            applyFilter();
        } catch (SQLException ex) {
            showDbError("Ошибка загрузки персонала", ex);
        }
    }

    private void applyFilter() {
        String query = searchField.getText() == null ? "" : searchField.getText().trim().toLowerCase();
        List<PersonnelRecord> filtered = allPersonnel;
        if (!query.isBlank()) {
            filtered = allPersonnel.stream()
                    .filter(p -> p.getFullName().toLowerCase().contains(query)
                            || String.valueOf(p.getId()).contains(query))
                    .collect(Collectors.toList());
        }
        personnelModel.setData(filtered);
        if (!filtered.isEmpty()) {
            personnelTable.setRowSelectionInterval(0, 0);
        } else {
            refreshUnavailabilityPanel(null);
        }
    }

    private void refreshUnavailabilityPanel(PersonnelRecord selected) {
        updateYearFilterOptions(selected);
        if (selected == null) {
            unavailabilityModel.setData(List.of());
            return;
        }
        applyUnavailabilityFilter(selected);
    }

    private void updateYearFilterOptions(PersonnelRecord selected) {
        Object previous = yearFilterCombo.getSelectedItem();
        yearFilterCombo.removeAllItems();
        yearFilterCombo.addItem(YEAR_FILTER_ALL);

        if (selected != null) {
            LinkedHashSet<Integer> years = selected.getUnavailabilityDates().stream()
                    .map(PersonnelRecord.UnavailabilityRecord::getUnavailableDate)
                    .map(PersonnelTab::parseRawDate)
                    .filter(Objects::nonNull)
                    .map(LocalDate::getYear)
                    .sorted()
                    .collect(Collectors.toCollection(LinkedHashSet::new));
            years.forEach(y -> yearFilterCombo.addItem(String.valueOf(y)));
        }

        if (previous != null) {
            yearFilterCombo.setSelectedItem(previous);
        }
        if (yearFilterCombo.getSelectedIndex() < 0) {
            yearFilterCombo.setSelectedIndex(0);
        }
    }

    private void applyUnavailabilityFilter(PersonnelRecord selected) {
        String selectedYearValue = Objects.toString(yearFilterCombo.getSelectedItem(), YEAR_FILTER_ALL);
        List<PersonnelRecord.UnavailabilityRecord> filtered = selected.getUnavailabilityDates().stream()
                .filter(rec -> {
                    if (YEAR_FILTER_ALL.equals(selectedYearValue)) {
                        return true;
                    }
                    LocalDate date = parseRawDate(rec.getUnavailableDate());
                    return date != null && String.valueOf(date.getYear()).equals(selectedYearValue);
                })
                .collect(Collectors.toList());
        unavailabilityModel.setData(groupUnavailability(filtered));
    }

    private static List<GroupedUnavailabilityRecord> groupUnavailability(List<PersonnelRecord.UnavailabilityRecord> records) {
        List<PersonnelRecord.UnavailabilityRecord> sortedRecords = records.stream()
                .sorted(Comparator.comparing(PersonnelRecord.UnavailabilityRecord::getUnavailableDate)
                        .thenComparing(r -> r.getReason() == null ? "" : r.getReason()))
                .toList();

        List<GroupedUnavailabilityRecord> grouped = new ArrayList<>();
        GroupedUnavailabilityRecord current = null;

        for (PersonnelRecord.UnavailabilityRecord record : sortedRecords) {
            LocalDate date = parseRawDate(record.getUnavailableDate());
            if (date == null) {
                continue;
            }
            String reason = record.getReason() == null ? "" : record.getReason().trim();
            if (current == null
                    || !current.reason.equals(reason)
                    || !current.toDate.plusDays(1).equals(date)) {
                current = new GroupedUnavailabilityRecord(new ArrayList<>(), date, date, reason);
                grouped.add(current);
            } else {
                current.toDate = date;
            }
            current.recordIds.add(record.getId());
        }

        return grouped;
    }

    private static LocalDate parseRawDate(String rawDate) {
        if (rawDate == null || rawDate.isBlank()) {
            return null;
        }
        String trimmed = rawDate.trim();
        try {
            return LocalDate.parse(trimmed);
        } catch (DateTimeParseException ignored) {
            try {
                return LocalDate.parse(trimmed, UI_DATE_FORMAT);
            } catch (DateTimeParseException ignoredAgain) {
                return null;
            }
        }
    }

    private PersonnelRecord getSelectedPerson() {
        int viewIndex = personnelTable.getSelectedRow();
        if (viewIndex < 0) return null;
        return personnelModel.getAt(viewIndex);
    }

    private void addPerson() {
        PersonForm form = PersonForm.showDialog(this, "Добавить сотрудника", null);
        if (form == null) return;
        try {
            DatabaseManager.addPersonnel(form.firstName, form.lastName, form.middleName);
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка добавления сотрудника", ex);
        }
    }

    private void editPerson() {
        PersonnelRecord selected = getSelectedPerson();
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Выберите сотрудника");
            return;
        }
        PersonForm form = PersonForm.showDialog(this, "Редактировать сотрудника", selected);
        if (form == null) return;
        try {
            DatabaseManager.updatePersonnel(selected.getId(), form.firstName, form.lastName, form.middleName);
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка редактирования сотрудника", ex);
        }
    }

    private void deletePerson() {
        PersonnelRecord selected = getSelectedPerson();
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Выберите сотрудника");
            return;
        }
        int confirm = JOptionPane.showConfirmDialog(this,
                "Удалить сотрудника " + selected.getFullName() + "?",
                "Подтверждение", JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) return;

        try {
            DatabaseManager.deletePersonnel(selected.getId());
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка удаления сотрудника", ex);
        }
    }

    private void addUnavailableDate() {
        PersonnelRecord selected = getSelectedPerson();
        if (selected == null) {
            JOptionPane.showMessageDialog(this, "Сначала выберите сотрудника");
            return;
        }

        JTextField fromDateField = new JTextField();
        fromDateField.setToolTipText("Формат даты: " + DATE_PLACEHOLDER);
        JTextField toDateField = new JTextField();
        toDateField.setToolTipText("Формат даты: " + DATE_PLACEHOLDER);
        JTextField reasonField = new JTextField();

        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        JLabel personLabel = new JLabel(selected.getFullName(), SwingConstants.CENTER);
        panel.add(personLabel);
        panel.add(new JLabel("Недоступен с:"));
        panel.add(fromDateField);
        panel.add(new JLabel("Формат: " + DATE_PLACEHOLDER));
        panel.add(new JLabel("Недоступен по:"));
        panel.add(toDateField);
        panel.add(new JLabel("Формат: " + DATE_PLACEHOLDER));
        panel.add(new JLabel("Причина (отпуск, аудит и т.д.):"));
        panel.add(reasonField);

        int result = JOptionPane.showConfirmDialog(this, panel,
                "Добавить период занятости", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result != JOptionPane.OK_OPTION) return;

        LocalDate from = parseUiDate(fromDateField.getText());
        LocalDate to = parseUiDate(toDateField.getText());
        if (from == null || to == null) {
            JOptionPane.showMessageDialog(this,
                    "Введите даты в формате " + DATE_PLACEHOLDER,
                    "Некорректный формат даты",
                    JOptionPane.WARNING_MESSAGE);
            return;
        }

        if (to.isBefore(from)) {
            JOptionPane.showMessageDialog(this, "Дата по не может быть раньше даты с");
            return;
        }

        String reason = reasonField.getText().trim();

        try {
            LocalDate cursor = from;
            while (!cursor.isAfter(to)) {
                DatabaseManager.addPersonnelUnavailability(selected.getId(), cursor.toString(), reason);
                cursor = cursor.plusDays(1);
            }
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка сохранения даты недоступности", ex);
        }
    }

    private LocalDate parseUiDate(String dateText) {
        if (dateText == null || dateText.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(dateText.trim(), UI_DATE_FORMAT);
        } catch (DateTimeParseException ex) {
            return null;
        }
    }

    private void deleteUnavailableDate() {
        PersonnelRecord selected = getSelectedPerson();
        int index = unavailabilityTable.getSelectedRow();
        if (selected == null || index < 0) {
            JOptionPane.showMessageDialog(this, "Выберите дату в нижней таблице");
            return;
        }
        GroupedUnavailabilityRecord rec = unavailabilityModel.getAt(index);

        int confirm = JOptionPane.showConfirmDialog(this,
                "Удалить выбранную дату занятости?",
                "Подтверждение удаления",
                JOptionPane.YES_NO_OPTION);
        if (confirm != JOptionPane.YES_OPTION) {
            return;
        }

        try {
            for (Integer recordId : rec.recordIds) {
                DatabaseManager.deletePersonnelUnavailability(recordId);
            }
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка удаления даты", ex);
        }
    }

    private void showDbError(String title, Exception ex) {
        JOptionPane.showMessageDialog(this, title + ": " + ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
    }

    private static class PersonForm {
        final String firstName;
        final String lastName;
        final String middleName;

        private PersonForm(String firstName, String lastName, String middleName) {
            this.firstName = firstName;
            this.lastName = lastName;
            this.middleName = middleName;
        }

        static PersonForm showDialog(Component parent, String title, PersonnelRecord current) {
            JTextField lastName = new JTextField(current == null ? "" : current.getLastName());
            JTextField firstName = new JTextField(current == null ? "" : current.getFirstName());
            JTextField middleName = new JTextField(current == null ? "" : current.getMiddleName());

            JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
            panel.add(new JLabel("Фамилия:"));
            panel.add(lastName);
            panel.add(new JLabel("Имя:"));
            panel.add(firstName);
            panel.add(new JLabel("Отчество:"));
            panel.add(middleName);

            int result = JOptionPane.showConfirmDialog(parent, panel, title,
                    JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
            if (result != JOptionPane.OK_OPTION) return null;
            return new PersonForm(firstName.getText().trim(), lastName.getText().trim(), middleName.getText().trim());
        }
    }

    private static class PersonnelTableModel extends AbstractTableModel {
        private final String[] columns = {"ID", "Фамилия", "Имя", "Отчество"};
        private List<PersonnelRecord> data = List.of();

        void setData(List<PersonnelRecord> data) {
            this.data = new ArrayList<>(data);
            fireTableDataChanged();
        }

        PersonnelRecord getAt(int row) {
            return data.get(row);
        }

        @Override public int getRowCount() { return data.size(); }
        @Override public int getColumnCount() { return columns.length; }
        @Override public String getColumnName(int col) { return columns[col]; }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            PersonnelRecord p = data.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> p.getId();
                case 1 -> p.getLastName();
                case 2 -> p.getFirstName();
                case 3 -> p.getMiddleName();
                default -> "";
            };
        }
    }

    private static class UnavailabilityTableModel extends AbstractTableModel {
        private final String[] columns = {"Дата недоступности", "Причина"};
        private List<GroupedUnavailabilityRecord> data = List.of();

        void setData(List<GroupedUnavailabilityRecord> data) {
            this.data = new ArrayList<>(data);
            fireTableDataChanged();
        }

        GroupedUnavailabilityRecord getAt(int row) {
            return data.get(row);
        }

        @Override public int getRowCount() { return data.size(); }
        @Override public int getColumnCount() { return columns.length; }
        @Override public String getColumnName(int col) { return columns[col]; }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            GroupedUnavailabilityRecord r = data.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> formatDateRangeForUi(r.fromDate, r.toDate);
                case 1 -> r.reason;
                default -> "";
            };
        }
    }

    private static String formatDateRangeForUi(LocalDate from, LocalDate to) {
        if (from == null || to == null) {
            return "";
        }
        if (from.equals(to)) {
            return from.format(UI_DATE_FORMAT);
        }
        return from.format(UI_DATE_FORMAT) + " - " + to.format(UI_DATE_FORMAT);
    }

    private static class GroupedUnavailabilityRecord {
        private final List<Integer> recordIds;
        private final LocalDate fromDate;
        private LocalDate toDate;
        private final String reason;

        private GroupedUnavailabilityRecord(List<Integer> recordIds, LocalDate fromDate, LocalDate toDate, String reason) {
            this.recordIds = recordIds;
            this.fromDate = fromDate;
            this.toDate = toDate;
            this.reason = reason;
        }
    }
}
