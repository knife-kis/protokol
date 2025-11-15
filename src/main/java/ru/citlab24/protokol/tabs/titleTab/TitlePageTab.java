package ru.citlab24.protokol.tabs.titleTab;

import com.github.lgooddatepicker.components.DatePicker;
import com.github.lgooddatepicker.components.DatePickerSettings;
import ru.citlab24.protokol.tabs.models.Building;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.FocusAdapter;
import java.awt.event.FocusEvent;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Вкладка «Титульная страница».
 * Основные реквизиты протокола, заказчика и условия проведения измерений.
 */
public class TitlePageTab extends JPanel {

    private final Building building;

    private static final String DATE_PATTERN = "dd.MM.yyyy";
    private static final DateTimeFormatter DATE_FORMAT =
            DateTimeFormatter.ofPattern(DATE_PATTERN);
    private static final String CUSTOMER_PLACEHOLDER = "Название, почта, телефон";

    // 1) Дата протокола (календарь)
    private final DatePicker protocolDatePicker;

    // 2) Наименование и контакты заказчика (с плейсхолдером)
    private final JTextField customerNameContactsField;

    // 3) Юридический адрес заказчика
    private final JTextField customerLegalAddressField;
    // 4) Фактический адрес заказчика
    private final JTextField customerActualAddressField;

    // 5) Наименование объекта измерений (многострочное)
    private final JTextArea objectNameArea;
    // 6) Адрес предприятия (объекта)
    private final JTextField objectAddressField;

    // 7) Договор: номер + дата (календарь)
    private final JTextField contractNumberField;
    private final DatePicker contractDatePicker;

    // 8) Заявка: номер + дата (календарь)
    private final JTextField applicationNumberField;
    private final DatePicker applicationDatePicker;

    // 9) Представитель заказчика (ФИО, должность)
    private final JTextField representativeField;

    // 10) Даты измерений и температуры
    private final JPanel measurementRowsPanel;
    private final List<MeasurementRow> measurementRows = new ArrayList<>();

    /** Одна строка измерений: панель, дата + 4 температуры. */
    private static class MeasurementRow {
        JPanel panel;
        DatePicker datePicker;
        JTextField tempInsideStart;
        JTextField tempInsideEnd;
        JTextField tempOutsideStart;
        JTextField tempOutsideEnd;
    }

    /** DTO для чтения данных измерений наружу. */
    public static class MeasurementRowData {
        public String date;             // dd.MM.yyyy
        public String tempInsideStart;  // начало внутри
        public String tempInsideEnd;    // конец внутри
        public String tempOutsideStart; // начало улица
        public String tempOutsideEnd;   // конец улица
    }

    public TitlePageTab(Building building) {
        this.building = building;

        setLayout(new BorderLayout());
        setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // ===== Верхняя часть — реквизиты протокола и заказчика =====
        JPanel formPanel = new JPanel(new GridBagLayout());
        formPanel.setBorder(new TitledBorder("Реквизиты протокола и заказчика"));

        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(4, 4, 4, 4);
        gbc.anchor = GridBagConstraints.WEST;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.weightx = 0.0;
        gbc.gridy = 0;

        // 1) Дата протокола (по умолчанию завтра, через календарь)
        protocolDatePicker = createDatePicker(LocalDate.now().plusDays(1));
        addLabeledField(formPanel, gbc, "Дата протокола:", protocolDatePicker);

        // 2) Наименование и контактные данные заказчика (с плейсхолдером)
        customerNameContactsField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Наименование и контактные данные заказчика:",
                customerNameContactsField);
        installPlaceholder(customerNameContactsField, CUSTOMER_PLACEHOLDER);

        // 3) Юридический адрес заказчика
        customerLegalAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Юридический адрес заказчика:",
                customerLegalAddressField);

        // 4) Фактический адрес заказчика
        customerActualAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Фактический адрес заказчика:",
                customerActualAddressField);

        // 5) Наименование объекта измерений — многострочное поле
        objectNameArea = new JTextArea(3, 40);
        objectNameArea.setLineWrap(true);
        objectNameArea.setWrapStyleWord(true);
        JScrollPane objectNameScroll = new JScrollPane(
                objectNameArea,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER
        );
        addLabeledField(formPanel, gbc,
                "Наименование объекта измерений:",
                objectNameScroll);

        // 6) Адрес предприятия (объекта)
        objectAddressField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Адрес предприятия (объекта):",
                objectAddressField);

        // 7) Договор: номер + дата (календарь)
        contractNumberField = new JTextField(10);
        contractDatePicker = createDatePicker(null);
        JPanel contractPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        contractPanel.add(new JLabel("№"));
        contractPanel.add(contractNumberField);
        contractPanel.add(new JLabel("от"));
        contractPanel.add(contractDatePicker);
        addLabeledField(formPanel, gbc,
                "Договор:",
                contractPanel);

        // 8) Заявка: номер + дата (календарь)
        applicationNumberField = new JTextField(10);
        applicationDatePicker = createDatePicker(null);
        JPanel applicationPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        applicationPanel.add(new JLabel("№"));
        applicationPanel.add(applicationNumberField);
        applicationPanel.add(new JLabel("от"));
        applicationPanel.add(applicationDatePicker);
        addLabeledField(formPanel, gbc,
                "Заявка:",
                applicationPanel);

        // 9) Представитель заказчика: ФИО и должность
        representativeField = new JTextField(40);
        addLabeledField(formPanel, gbc,
                "Представитель заказчика (ФИО, должность):",
                representativeField);

        // ===== Нижняя часть — даты измерений и температуры =====
        JPanel bottomPanel = new JPanel(new BorderLayout());
        bottomPanel.setBorder(new TitledBorder("Даты проведения измерений и температуры воздуха"));

        // Панель строк "дата + 4 температуры"
        measurementRowsPanel = new JPanel();
        measurementRowsPanel.setLayout(new BoxLayout(measurementRowsPanel, BoxLayout.Y_AXIS));
        JScrollPane measurementScroll = new JScrollPane(
                measurementRowsPanel,
                JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED,
                JScrollPane.HORIZONTAL_SCROLLBAR_NEVER
        );

        // Кнопки "Добавить дату" / "Удалить последнюю"
        JPanel buttonsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 0));
        JButton addRowButton = new JButton("Добавить дату измерений");
        JButton removeRowButton = new JButton("Удалить последнюю строку");

        addRowButton.addActionListener(e -> addMeasurementRow(null));
        removeRowButton.addActionListener(e -> removeLastMeasurementRow());

        buttonsPanel.add(addRowButton);
        buttonsPanel.add(removeRowButton);

        // По умолчанию — одна строка измерений
        addMeasurementRow(null);

        JPanel topBottomPanel = new JPanel(new BorderLayout());
        topBottomPanel.add(buttonsPanel, BorderLayout.NORTH);
        topBottomPanel.add(measurementScroll, BorderLayout.CENTER);

        bottomPanel.add(topBottomPanel, BorderLayout.CENTER);

        // Итог: верхняя и нижняя части на вкладке
        add(formPanel, BorderLayout.NORTH);
        add(bottomPanel, BorderLayout.CENTER);
    }

    /* ================= Вспомогательные методы UI ================= */

    private void addLabeledField(JPanel panel, GridBagConstraints gbc,
                                 String labelText, JComponent field) {
        gbc.gridx = 0;
        gbc.weightx = 0.0;
        gbc.fill = GridBagConstraints.NONE;
        JLabel label = new JLabel(labelText);
        panel.add(label, gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(field, gbc);

        gbc.gridy++;
    }

    private DatePicker createDatePicker(LocalDate initialDate) {
        DatePickerSettings settings = new DatePickerSettings(new Locale("ru", "RU"));
        settings.setFormatForDatesCommonEra(DATE_PATTERN);
        settings.setAllowEmptyDates(true);
        DatePicker picker = new DatePicker(settings);
        if (initialDate != null) {
            picker.setDate(initialDate);
        }
        return picker;
    }

    /**
     * Плейсхолдер: серый текст, который исчезает при вводе и
     * не возвращается из геттера.
     */
    private void installPlaceholder(JTextField field, String placeholder) {
        Color normalColor = field.getForeground();
        field.setForeground(Color.GRAY);
        field.setText(placeholder);

        field.addFocusListener(new FocusAdapter() {
            @Override
            public void focusGained(FocusEvent e) {
                if (field.getText().equals(placeholder)) {
                    field.setText("");
                    field.setForeground(normalColor);
                }
            }

            @Override
            public void focusLost(FocusEvent e) {
                if (field.getText().isBlank()) {
                    field.setForeground(Color.GRAY);
                    field.setText(placeholder);
                }
            }
        });
    }

    private void addMeasurementRow(LocalDate date) {
        MeasurementRow row = new MeasurementRow();
        row.panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 4, 2));

        int index = measurementRows.size() + 1;

        row.datePicker = createDatePicker(date);
        row.tempInsideStart = new JTextField(4);
        row.tempInsideEnd   = new JTextField(4);
        row.tempOutsideStart = new JTextField(4);
        row.tempOutsideEnd   = new JTextField(4);

        row.panel.add(new JLabel("Дата " + index + ":"));
        row.panel.add(row.datePicker);

        row.panel.add(new JLabel("Внутри, начало:"));
        row.panel.add(row.tempInsideStart);
        row.panel.add(new JLabel("конец:"));
        row.panel.add(row.tempInsideEnd);

        row.panel.add(new JLabel("Улица, начало:"));
        row.panel.add(row.tempOutsideStart);
        row.panel.add(new JLabel("конец:"));
        row.panel.add(row.tempOutsideEnd);

        measurementRows.add(row);
        measurementRowsPanel.add(row.panel);
        measurementRowsPanel.revalidate();
        measurementRowsPanel.repaint();
    }

    private void removeLastMeasurementRow() {
        if (measurementRows.isEmpty()) return;
        MeasurementRow last = measurementRows.remove(measurementRows.size() - 1);
        measurementRowsPanel.remove(last.panel);
        measurementRowsPanel.revalidate();
        measurementRowsPanel.repaint();
    }

    /* ================= Геттеры для сохранения/экспорта ================= */

    // Дата протокола
    public String getProtocolDate() {
        LocalDate d = protocolDatePicker.getDate();
        return (d != null) ? d.format(DATE_FORMAT) : "";
    }

    // Наименование и контакты заказчика
    public String getCustomerNameAndContacts() {
        String text = customerNameContactsField.getText().trim();
        if (text.equals(CUSTOMER_PLACEHOLDER)) {
            return "";
        }
        return text;
    }

    public String getCustomerLegalAddress() {
        return customerLegalAddressField.getText().trim();
    }

    public String getCustomerActualAddress() {
        return customerActualAddressField.getText().trim();
    }

    public String getObjectName() {
        return objectNameArea.getText().trim();
    }

    public String getObjectAddress() {
        return objectAddressField.getText().trim();
    }

    // Договор
    public String getContractNumber() {
        return contractNumberField.getText().trim();
    }

    public String getContractDate() {
        LocalDate d = contractDatePicker.getDate();
        return (d != null) ? d.format(DATE_FORMAT) : "";
    }

    public String getContractInfo() {
        String num = getContractNumber();
        String date = getContractDate();
        if (num.isEmpty() && date.isEmpty()) return "";
        if (date.isEmpty()) return "№ " + num;
        if (num.isEmpty()) return "от " + date;
        return "№ " + num + " от " + date;
    }

    // Заявка
    public String getApplicationNumber() {
        return applicationNumberField.getText().trim();
    }

    public String getApplicationDate() {
        LocalDate d = applicationDatePicker.getDate();
        return (d != null) ? d.format(DATE_FORMAT) : "";
    }

    public String getApplicationInfo() {
        String num = getApplicationNumber();
        String date = getApplicationDate();
        if (num.isEmpty() && date.isEmpty()) return "";
        if (date.isEmpty()) return "№ " + num;
        if (num.isEmpty()) return "от " + date;
        return "№ " + num + " от " + date;
    }

    // Представитель заказчика
    public String getRepresentative() {
        return representativeField.getText().trim();
    }

    /** Все строки «дата измерений + 4 температуры». Пустые строки пропускаем. */
    public List<MeasurementRowData> getMeasurementRows() {
        List<MeasurementRowData> list = new ArrayList<>();
        for (MeasurementRow row : measurementRows) {
            LocalDate d = row.datePicker.getDate();
            String inStart = row.tempInsideStart.getText().trim();
            String inEnd   = row.tempInsideEnd.getText().trim();
            String outStart = row.tempOutsideStart.getText().trim();
            String outEnd   = row.tempOutsideEnd.getText().trim();

            boolean allEmpty =
                    (d == null) &&
                            inStart.isEmpty() &&
                            inEnd.isEmpty() &&
                            outStart.isEmpty() &&
                            outEnd.isEmpty();

            if (allEmpty) continue;

            MeasurementRowData data = new MeasurementRowData();
            data.date = (d != null) ? d.format(DATE_FORMAT) : "";
            data.tempInsideStart = inStart;
            data.tempInsideEnd = inEnd;
            data.tempOutsideStart = outStart;
            data.tempOutsideEnd = outEnd;
            list.add(data);
        }
        return list;
    }

    public Building getBuilding() {
        return building;
    }
}
