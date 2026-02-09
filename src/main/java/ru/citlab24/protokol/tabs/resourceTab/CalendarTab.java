package ru.citlab24.protokol.tabs.resourceTab;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.PersonnelRecord;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import java.awt.*;
import java.sql.SQLException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.temporal.TemporalAdjusters;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

public class CalendarTab extends JPanel {
    private static final DateTimeFormatter MONTH_LABEL = DateTimeFormatter.ofPattern("LLLL yyyy", new Locale("ru"));
    private static final DateTimeFormatter DAY_LABEL = DateTimeFormatter.ofPattern("EEEE, dd MMMM yyyy", new Locale("ru"));

    private static final String ASPECT_VACATION = "Отпуск";
    private static final String ASPECT_AUDIT = "Аудит";
    private static final String ASPECT_TESTS = "Испытания";
    private static final String ASPECT_MSI = "МСИ";
    private static final String ASPECT_VERIFICATION = "Поверка оборудования";

    private final JCheckBox showAllCheck = new JCheckBox("Показать все");
    private final Map<String, JCheckBox> aspectFilters = new HashMap<>();

    private final JButton prevButton = new JButton("◀");
    private final JButton nextButton = new JButton("▶");
    private final JButton todayButton = new JButton("Сегодня");
    private final JToggleButton yearScaleToggle = new JToggleButton("Год (мелкий масштаб)");
    private final JLabel monthTitle = new JLabel("", SwingConstants.CENTER);

    private final JPanel calendarContainer = new JPanel(new BorderLayout());
    private final JLabel selectedDateLabel = new JLabel("Выберите день");
    private final JTextArea detailsArea = new JTextArea();

    private YearMonth currentMonth = YearMonth.now();
    private LocalDate selectedDate = LocalDate.now();
    private final List<CalendarEvent> allEvents = new ArrayList<>();

    public CalendarTab() {
        super(new BorderLayout(8, 8));
        add(createToolbar(), BorderLayout.NORTH);
        add(createCenterContent(), BorderLayout.CENTER);
        reloadEvents();
        renderCalendar();
    }

    private JComponent createToolbar() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));

        JPanel navigation = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        navigation.add(prevButton);
        navigation.add(nextButton);
        navigation.add(todayButton);
        navigation.add(yearScaleToggle);

        prevButton.addActionListener(e -> {
            currentMonth = yearScaleToggle.isSelected() ? currentMonth.minusYears(1) : currentMonth.minusMonths(1);
            renderCalendar();
        });
        nextButton.addActionListener(e -> {
            currentMonth = yearScaleToggle.isSelected() ? currentMonth.plusYears(1) : currentMonth.plusMonths(1);
            renderCalendar();
        });
        todayButton.addActionListener(e -> {
            currentMonth = YearMonth.now();
            selectedDate = LocalDate.now();
            renderCalendar();
            showSelectedDayDetails();
        });
        yearScaleToggle.addActionListener(e -> renderCalendar());

        monthTitle.setFont(monthTitle.getFont().deriveFont(Font.BOLD, 18f));

        panel.add(navigation, BorderLayout.WEST);
        panel.add(monthTitle, BorderLayout.CENTER);
        panel.add(createFilterPanel(), BorderLayout.EAST);
        return panel;
    }

    private JComponent createFilterPanel() {
        JPanel panel = new JPanel(new GridLayout(0, 1, 0, 2));
        panel.setBorder(new EmptyBorder(0, 8, 0, 0));

        showAllCheck.setSelected(true);
        showAllCheck.addActionListener(e -> {
            boolean enabled = !showAllCheck.isSelected();
            for (JCheckBox box : aspectFilters.values()) {
                box.setEnabled(enabled);
            }
            renderCalendar();
            showSelectedDayDetails();
        });
        panel.add(showAllCheck);

        addAspectCheckbox(panel, ASPECT_VACATION);
        addAspectCheckbox(panel, ASPECT_AUDIT);
        addAspectCheckbox(panel, ASPECT_TESTS);
        addAspectCheckbox(panel, ASPECT_MSI);
        addAspectCheckbox(panel, ASPECT_VERIFICATION);

        for (JCheckBox box : aspectFilters.values()) {
            box.setEnabled(false);
        }

        return panel;
    }

    private void addAspectCheckbox(JPanel panel, String aspect) {
        JCheckBox box = new JCheckBox(aspect, true);
        box.addActionListener(e -> {
            renderCalendar();
            showSelectedDayDetails();
        });
        aspectFilters.put(aspect, box);
        panel.add(box);
    }

    private JComponent createCenterContent() {
        detailsArea.setEditable(false);
        detailsArea.setLineWrap(true);
        detailsArea.setWrapStyleWord(true);

        JPanel detailsPanel = new JPanel(new BorderLayout(6, 6));
        detailsPanel.setBorder(BorderFactory.createTitledBorder("Подробности дня"));
        detailsPanel.add(selectedDateLabel, BorderLayout.NORTH);
        detailsPanel.add(new JScrollPane(detailsArea), BorderLayout.CENTER);

        JSplitPane splitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, calendarContainer, detailsPanel);
        splitPane.setResizeWeight(0.75);
        return splitPane;
    }

    private void reloadEvents() {
        allEvents.clear();
        try {
            for (PersonnelRecord person : DatabaseManager.getAllPersonnel()) {
                for (PersonnelRecord.UnavailabilityRecord rec : person.getUnavailabilityDates()) {
                    LocalDate day = parseDate(rec.getUnavailableDate());
                    if (day == null) {
                        continue;
                    }
                    String aspect = normalizeAspect(rec.getReason());
                    String shortName = person.getLastName() == null ? person.getFullName() : person.getLastName();
                    String title = aspect.equals(ASPECT_VACATION)
                            ? "Отпуск " + shortName
                            : aspect + " " + shortName;
                    allEvents.add(new CalendarEvent(day, aspect, title, person.getFullName()));
                }
            }
            allEvents.sort(Comparator.comparing(CalendarEvent::date).thenComparing(CalendarEvent::title));
        } catch (SQLException ex) {
            JOptionPane.showMessageDialog(this,
                    "Не удалось загрузить календарь: " + ex.getMessage(),
                    "Ошибка БД",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private void renderCalendar() {
        calendarContainer.removeAll();
        monthTitle.setText(yearScaleToggle.isSelected()
                ? String.valueOf(currentMonth.getYear())
                : capitalize(currentMonth.format(MONTH_LABEL)));

        if (yearScaleToggle.isSelected()) {
            calendarContainer.add(createYearView(currentMonth.getYear()), BorderLayout.CENTER);
        } else {
            calendarContainer.add(createMonthView(currentMonth), BorderLayout.CENTER);
        }

        calendarContainer.revalidate();
        calendarContainer.repaint();
        showSelectedDayDetails();
    }

    private JComponent createMonthView(YearMonth month) {
        JPanel root = new JPanel(new BorderLayout(4, 4));
        root.add(createWeekHeader(), BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(0, 7, 4, 4));

        LocalDate firstDay = month.atDay(1);
        LocalDate start = firstDay.with(TemporalAdjusters.previousOrSame(DayOfWeek.MONDAY));
        LocalDate end = month.atEndOfMonth().with(TemporalAdjusters.nextOrSame(DayOfWeek.SUNDAY));

        for (LocalDate d = start; !d.isAfter(end); d = d.plusDays(1)) {
            grid.add(createDayCell(d, month, false));
        }

        root.add(new JScrollPane(grid), BorderLayout.CENTER);
        return root;
    }

    private JComponent createYearView(int year) {
        JPanel monthsGrid = new JPanel(new GridLayout(4, 3, 8, 8));
        for (int m = 1; m <= 12; m++) {
            YearMonth ym = YearMonth.of(year, m);
            JPanel panel = new JPanel(new BorderLayout(2, 2));
            JLabel label = new JLabel(capitalize(ym.format(DateTimeFormatter.ofPattern("LLLL", new Locale("ru")))), SwingConstants.CENTER);
            label.setFont(label.getFont().deriveFont(Font.BOLD, 12f));
            panel.add(label, BorderLayout.NORTH);

            JPanel mini = new JPanel(new GridLayout(0, 7, 1, 1));
            LocalDate firstDay = ym.atDay(1);
            LocalDate start = firstDay.with(TemporalAdjusters.previousOrSame(DayOfWeek.MONDAY));
            LocalDate end = ym.atEndOfMonth().with(TemporalAdjusters.nextOrSame(DayOfWeek.SUNDAY));

            for (LocalDate d = start; !d.isAfter(end); d = d.plusDays(1)) {
                mini.add(createDayCell(d, ym, true));
            }

            panel.add(mini, BorderLayout.CENTER);
            panel.setBorder(BorderFactory.createLineBorder(new Color(220, 220, 220)));
            monthsGrid.add(panel);
        }
        return new JScrollPane(monthsGrid);
    }

    private JComponent createWeekHeader() {
        JPanel header = new JPanel(new GridLayout(1, 7, 4, 4));
        String[] days = {"Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"};
        for (String day : days) {
            JLabel label = new JLabel(day, SwingConstants.CENTER);
            label.setFont(label.getFont().deriveFont(Font.BOLD));
            header.add(label);
        }
        return header;
    }

    private JComponent createDayCell(LocalDate date, YearMonth visibleMonth, boolean compact) {
        List<CalendarEvent> events = eventsForDate(date);

        JPanel cell = new JPanel(new BorderLayout(2, 2));
        cell.setBorder(BorderFactory.createLineBorder(new Color(210, 210, 210)));
        cell.setBackground(Color.WHITE);

        if (!date.getMonth().equals(visibleMonth.getMonth())) {
            cell.setBackground(new Color(248, 248, 248));
        }
        if (date.equals(LocalDate.now())) {
            cell.setBorder(BorderFactory.createLineBorder(new Color(66, 133, 244), 2));
        }
        if (date.equals(selectedDate)) {
            cell.setBackground(new Color(232, 242, 255));
        }

        JLabel dayNumber = new JLabel(String.valueOf(date.getDayOfMonth()));
        dayNumber.setBorder(new EmptyBorder(2, 4, 0, 0));
        dayNumber.setFont(dayNumber.getFont().deriveFont(compact ? 9f : 11f));
        cell.add(dayNumber, BorderLayout.NORTH);

        if (!compact) {
            JPanel eventsPanel = new JPanel();
            eventsPanel.setLayout(new BoxLayout(eventsPanel, BoxLayout.Y_AXIS));
            eventsPanel.setOpaque(false);

            int limit = 3;
            for (int i = 0; i < Math.min(limit, events.size()); i++) {
                JLabel eLabel = new JLabel("• " + events.get(i).title());
                eLabel.setFont(eLabel.getFont().deriveFont(10f));
                eventsPanel.add(eLabel);
            }
            if (events.size() > limit) {
                JLabel more = new JLabel("+" + (events.size() - limit) + " ещё");
                more.setFont(more.getFont().deriveFont(Font.ITALIC, 10f));
                eventsPanel.add(more);
            }
            cell.add(eventsPanel, BorderLayout.CENTER);
        } else if (!events.isEmpty()) {
            JPanel marker = new JPanel();
            marker.setBackground(new Color(93, 173, 226));
            marker.setPreferredSize(new Dimension(8, 8));
            cell.add(marker, BorderLayout.SOUTH);
        }

        cell.setToolTipText(events.isEmpty() ? "Нет событий" : events.stream().map(CalendarEvent::title).collect(Collectors.joining(", ")));
        cell.addMouseListener(new java.awt.event.MouseAdapter() {
            @Override
            public void mouseClicked(java.awt.event.MouseEvent e) {
                selectedDate = date;
                if (yearScaleToggle.isSelected()) {
                    currentMonth = YearMonth.from(date);
                }
                renderCalendar();
                showSelectedDayDetails();
            }
        });

        return cell;
    }

    private void showSelectedDayDetails() {
        selectedDateLabel.setText(capitalize(selectedDate.format(DAY_LABEL)));
        List<CalendarEvent> events = eventsForDate(selectedDate);
        if (events.isEmpty()) {
            detailsArea.setText("На выбранную дату событий не найдено.");
            return;
        }

        StringBuilder text = new StringBuilder();
        for (CalendarEvent event : events) {
            text.append("• ")
                    .append(event.aspect())
                    .append(": ")
                    .append(event.title())
                    .append("\n  ")
                    .append(event.details())
                    .append("\n\n");
        }
        detailsArea.setText(text.toString().trim());
        detailsArea.setCaretPosition(0);
    }

    private List<CalendarEvent> eventsForDate(LocalDate date) {
        Set<String> selectedAspects = selectedAspects();
        return allEvents.stream()
                .filter(e -> e.date().equals(date))
                .filter(e -> selectedAspects.contains(e.aspect()))
                .toList();
    }

    private Set<String> selectedAspects() {
        if (showAllCheck.isSelected()) {
            return aspectFilters.keySet();
        }
        return aspectFilters.entrySet().stream()
                .filter(entry -> entry.getValue().isSelected())
                .map(Map.Entry::getKey)
                .collect(Collectors.toSet());
    }

    private static LocalDate parseDate(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(value.trim());
        } catch (DateTimeParseException ignored) {
            try {
                return LocalDate.parse(value.trim(), DateTimeFormatter.ofPattern("dd-MM-yyyy"));
            } catch (DateTimeParseException ignoredAgain) {
                return null;
            }
        }
    }

    private String normalizeAspect(String rawReason) {
        String reason = rawReason == null ? "" : rawReason.trim().toLowerCase(new Locale("ru"));
        if (reason.isBlank() || reason.contains("отпуск")) {
            return ASPECT_VACATION;
        }
        if (reason.contains("аудит")) {
            return ASPECT_AUDIT;
        }
        if (reason.contains("испыт")) {
            return ASPECT_TESTS;
        }
        if (reason.contains("мси")) {
            return ASPECT_MSI;
        }
        if (reason.contains("поверк")) {
            return ASPECT_VERIFICATION;
        }
        return ASPECT_VACATION;
    }

    private static String capitalize(String text) {
        if (text == null || text.isBlank()) {
            return "";
        }
        return Character.toUpperCase(text.charAt(0)) + text.substring(1);
    }

    private record CalendarEvent(LocalDate date, String aspect, String title, String details) {
    }
}
