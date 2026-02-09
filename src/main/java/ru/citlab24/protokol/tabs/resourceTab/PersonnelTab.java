package ru.citlab24.protokol.tabs.resourceTab;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.PersonnelRecord;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.text.DateFormatter;
import java.awt.*;
import java.sql.SQLException;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.stream.Collectors;

public class PersonnelTab extends JPanel {
    private static final DateTimeFormatter UI_DATE_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");

    private final PersonnelTableModel personnelModel = new PersonnelTableModel();
    private final JTable personnelTable = new JTable(personnelModel);

    private final UnavailabilityTableModel unavailabilityModel = new UnavailabilityTableModel();
    private final JTable unavailabilityTable = new JTable(unavailabilityModel);

    private final JTextField searchField = new JTextField();
    private final List<PersonnelRecord> allPersonnel = new ArrayList<>();

    public PersonnelTab() {
        super(new BorderLayout(8, 8));
        add(createToolbar(), BorderLayout.NORTH);
        add(createContent(), BorderLayout.CENTER);

        personnelTable.getSelectionModel().addListSelectionListener(e -> {
            if (!e.getValueIsAdjusting()) {
                PersonnelRecord selected = getSelectedPerson();
                unavailabilityModel.setData(selected == null ? List.of() : selected.getUnavailabilityDates());
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
        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
                new JScrollPane(personnelTable),
                new JScrollPane(unavailabilityTable));
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
            unavailabilityModel.setData(List.of());
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

        JSpinner fromDateSpinner = createDateSpinner(new Date());
        JSpinner toDateSpinner = createDateSpinner(new Date());
        JTextField reasonField = new JTextField();
        JPanel panel = new JPanel(new GridLayout(0, 1, 4, 4));
        panel.add(new JLabel("Недоступен с:"));
        panel.add(fromDateSpinner);
        panel.add(new JLabel("Недоступен по:"));
        panel.add(toDateSpinner);
        panel.add(new JLabel("Причина (отпуск, аудит и т.д.):"));
        panel.add(reasonField);

        int result = JOptionPane.showConfirmDialog(this, panel,
                "Добавить период занятости", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result != JOptionPane.OK_OPTION) return;

        LocalDate from = toLocalDate((Date) fromDateSpinner.getValue());
        LocalDate to = toLocalDate((Date) toDateSpinner.getValue());

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

    private JSpinner createDateSpinner(Date initialDate) {
        SpinnerDateModel model = new SpinnerDateModel(initialDate, null, null, java.util.Calendar.DAY_OF_MONTH);
        JSpinner spinner = new JSpinner(model);
        JSpinner.DateEditor editor = new JSpinner.DateEditor(spinner, "dd-MM-yyyy");
        DateFormatter formatter = (DateFormatter) editor.getTextField().getFormatter();
        formatter.setAllowsInvalid(false);
        formatter.setOverwriteMode(true);
        spinner.setEditor(editor);
        return spinner;
    }

    private LocalDate toLocalDate(Date date) {
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    private void deleteUnavailableDate() {
        PersonnelRecord selected = getSelectedPerson();
        int index = unavailabilityTable.getSelectedRow();
        if (selected == null || index < 0) {
            JOptionPane.showMessageDialog(this, "Выберите дату в нижней таблице");
            return;
        }
        PersonnelRecord.UnavailabilityRecord rec = unavailabilityModel.getAt(index);
        try {
            DatabaseManager.deletePersonnelUnavailability(rec.getId());
            reloadPersonnel();
        } catch (SQLException ex) {
            showDbError("Ошибка удаления даты", ex);
        }
    }

    private void showDbError(String title, Exception ex) {
        JOptionPane.showMessageDialog(this, title + ": " + ex.getMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
    }

    private String formatDateForUi(String rawDate) {
        if (rawDate == null || rawDate.isBlank()) {
            return "";
        }
        String trimmed = rawDate.trim();
        try {
            return LocalDate.parse(trimmed).format(UI_DATE_FORMAT);
        } catch (DateTimeParseException ignored) {
            try {
                return LocalDate.parse(trimmed, UI_DATE_FORMAT).format(UI_DATE_FORMAT);
            } catch (DateTimeParseException ignoredAgain) {
                return trimmed;
            }
        }
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
        private final String[] columns = {"ID", "Дата недоступности", "Причина"};
        private List<PersonnelRecord.UnavailabilityRecord> data = List.of();

        void setData(List<PersonnelRecord.UnavailabilityRecord> data) {
            this.data = new ArrayList<>(data);
            fireTableDataChanged();
        }

        PersonnelRecord.UnavailabilityRecord getAt(int row) {
            return data.get(row);
        }

        @Override public int getRowCount() { return data.size(); }
        @Override public int getColumnCount() { return columns.length; }
        @Override public String getColumnName(int col) { return columns[col]; }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            PersonnelRecord.UnavailabilityRecord r = data.get(rowIndex);
            return switch (columnIndex) {
                case 0 -> r.getId();
                case 1 -> formatDateForUi(r.getUnavailableDate());
                case 2 -> r.getReason();
                default -> "";
            };
        }
    }
}
