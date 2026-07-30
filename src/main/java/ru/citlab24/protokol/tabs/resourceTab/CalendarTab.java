package ru.citlab24.protokol.tabs.resourceTab;

import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.db.PersonnelRecord;
import ru.citlab24.protokol.db.VlkDateRecord;
import ru.citlab24.protokol.visits.SiteVisitRecord;
import ru.citlab24.protokol.visits.SiteVisitRepository;
import ru.citlab24.protokol.visits.SiteVisitStatus;

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
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.stream.Collectors;

public class CalendarTab extends JPanel {
    private static final DateTimeFormatter MONTH_LABEL = DateTimeFormatter.ofPattern("LLLL yyyy", new Locale("ru"));
    private static final DateTimeFormatter DAY_LABEL = DateTimeFormatter.ofPattern("EEEE, dd MMMM yyyy", new Locale("ru"));
    private static final DateTimeFormatter DOTTED_DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final DateTimeFormatter DASHED_DATE_FORMAT = DateTimeFormatter.ofPattern("dd-MM-yyyy");
    private static final DateTimeFormatter TIME_LABEL = DateTimeFormatter.ofPattern("HH:mm");

    private static final Color CALENDAR_BACKGROUND = new Color(245, 247, 250);
    private static final Color SURFACE_COLOR = Color.WHITE;
    private static final Color GRID_COLOR = new Color(224, 228, 233);
    private static final Color MUTED_TEXT_COLOR = new Color(125, 133, 143);
    private static final Color PRIMARY_COLOR = new Color(45, 124, 231);
    private static final Color SELECTED_DAY_COLOR = new Color(232, 242, 255);

    private static final String ASPECT_VACATION = "Отпуск";
    private static final String ASPECT_AUDIT = "Аудит";
    private static final String ASPECT_TESTS = "Испытания";
    private static final String ASPECT_MSI = "МСИ";
    private static final String ASPECT_VERIFICATION = "Поверка оборудования";
    private static final String ASPECT_VLK = "ВЛК";

    private static final Map<String, Color> ASPECT_COLORS = Map.of(
            ASPECT_VACATION, new Color(66, 133, 244),
            ASPECT_AUDIT, new Color(251, 188, 5),
            ASPECT_TESTS, new Color(52, 168, 83),
            ASPECT_MSI, new Color(156, 39, 176),
            ASPECT_VERIFICATION, new Color(0, 172, 193),
            ASPECT_VLK, new Color(255, 112, 67)
    );

    private final JCheckBox showAllCheck = new JCheckBox("Показать все");
    private final Map<String, JCheckBox> aspectFilters = new HashMap<>();

    private final JButton prevButton = new JButton("‹");
    private final JButton nextButton = new JButton("›");
    private final JButton todayButton = new JButton("Сегодня");
    private final JToggleButton yearScaleToggle = new JToggleButton("Год");
    private final JLabel monthTitle = new JLabel("", SwingConstants.LEFT);

    private final JPanel calendarContainer = new JPanel(new BorderLayout());
    private final JLabel selectedDateLabel = new JLabel("Выберите день");
    private final JPanel detailsEventsPanel = new JPanel();

    private YearMonth currentMonth = YearMonth.now();
    private LocalDate selectedDate = LocalDate.now();
    private final List<CalendarEvent> allEvents = new ArrayList<>();

    public CalendarTab() {
        super(new BorderLayout());
        setBackground(CALENDAR_BACKGROUND);
        add(createToolbar(), BorderLayout.NORTH);
        add(createCenterContent(), BorderLayout.CENTER);
        reloadEvents();
        renderCalendar();
    }

    private JComponent createToolbar() {
        JPanel panel = new JPanel(new BorderLayout(12, 0));
        panel.setBackground(SURFACE_COLOR);
        panel.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(0, 0, 1, 0, GRID_COLOR),
                new EmptyBorder(10, 14, 10, 14)
        ));

        JPanel navigation = new JPanel(new FlowLayout(FlowLayout.LEFT, 6, 0));
        navigation.setOpaque(false);
        navigation.add(todayButton);
        navigation.add(prevButton);
        navigation.add(nextButton);

        configureToolbarButton(todayButton, 92);
        configureToolbarButton(prevButton, 36);
        configureToolbarButton(nextButton, 36);
        configureToolbarButton(yearScaleToggle, 72);
        prevButton.setToolTipText("Предыдущий период");
        nextButton.setToolTipText("Следующий период");
        yearScaleToggle.setToolTipText("Показать весь год");

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

        monthTitle.setFont(monthTitle.getFont().deriveFont(Font.BOLD, 20f));
        monthTitle.setBorder(new EmptyBorder(0, 6, 0, 0));

        panel.add(navigation, BorderLayout.WEST);
        panel.add(monthTitle, BorderLayout.CENTER);
        panel.add(yearScaleToggle, BorderLayout.EAST);
        return panel;
    }

    private void configureToolbarButton(AbstractButton button, int width) {
        button.setPreferredSize(new Dimension(width, 32));
        button.setFocusable(false);
        button.putClientProperty("JButton.buttonType", "roundRect");
    }

    private JComponent createFilterPanel() {
        JPanel panel = new JPanel();
        panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
        panel.setOpaque(false);

        JLabel title = new JLabel("Календари");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 15f));
        title.setAlignmentX(Component.LEFT_ALIGNMENT);
        panel.add(title);
        panel.add(Box.createVerticalStrut(8));

        showAllCheck.setSelected(true);
        showAllCheck.setOpaque(false);
        showAllCheck.setAlignmentX(Component.LEFT_ALIGNMENT);
        showAllCheck.addActionListener(e -> {
            boolean enabled = !showAllCheck.isSelected();
            for (JCheckBox box : aspectFilters.values()) {
                box.setEnabled(enabled);
            }
            renderCalendar();
            showSelectedDayDetails();
        });
        panel.add(showAllCheck);
        panel.add(Box.createVerticalStrut(4));

        addAspectCheckbox(panel, ASPECT_VACATION);
        addAspectCheckbox(panel, ASPECT_AUDIT);
        addAspectCheckbox(panel, ASPECT_TESTS);
        addAspectCheckbox(panel, ASPECT_MSI);
        addAspectCheckbox(panel, ASPECT_VERIFICATION);
        addAspectCheckbox(panel, ASPECT_VLK);

        for (JCheckBox box : aspectFilters.values()) {
            box.setEnabled(false);
        }

        return panel;
    }

    private void addAspectCheckbox(JPanel panel, String aspect) {
        JCheckBox box = new JCheckBox(aspect, true);
        box.setOpaque(false);
        box.addActionListener(e -> {
            renderCalendar();
            showSelectedDayDetails();
        });
        aspectFilters.put(aspect, box);

        JLabel marker = new JLabel("●");
        marker.setForeground(colorForAspect(aspect));
        marker.setFont(marker.getFont().deriveFont(Font.PLAIN, 15f));
        marker.setPreferredSize(new Dimension(18, 22));

        JPanel row = new JPanel(new BorderLayout(4, 0));
        row.setOpaque(false);
        row.setAlignmentX(Component.LEFT_ALIGNMENT);
        row.setMaximumSize(new Dimension(Integer.MAX_VALUE, 26));
        row.add(marker, BorderLayout.WEST);
        row.add(box, BorderLayout.CENTER);
        panel.add(row);
    }

    private JComponent createCenterContent() {
        calendarContainer.setBackground(CALENDAR_BACKGROUND);

        detailsEventsPanel.setLayout(new BoxLayout(detailsEventsPanel, BoxLayout.Y_AXIS));
        detailsEventsPanel.setBackground(SURFACE_COLOR);

        selectedDateLabel.setFont(selectedDateLabel.getFont().deriveFont(Font.BOLD, 15f));

        JScrollPane detailsScroll = new JScrollPane(detailsEventsPanel);
        detailsScroll.setBorder(null);
        detailsScroll.getViewport().setBackground(SURFACE_COLOR);
        detailsScroll.setHorizontalScrollBarPolicy(ScrollPaneConstants.HORIZONTAL_SCROLLBAR_NEVER);
        detailsScroll.getVerticalScrollBar().setUnitIncrement(20);

        JPanel dayDetails = new JPanel(new BorderLayout(0, 10));
        dayDetails.setOpaque(false);
        dayDetails.add(selectedDateLabel, BorderLayout.NORTH);
        dayDetails.add(detailsScroll, BorderLayout.CENTER);

        JPanel sidebar = new JPanel(new BorderLayout(0, 18));
        sidebar.setBackground(SURFACE_COLOR);
        sidebar.setPreferredSize(new Dimension(292, 0));
        sidebar.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(0, 1, 0, 0, GRID_COLOR),
                new EmptyBorder(16, 16, 16, 16)
        ));
        sidebar.add(createFilterPanel(), BorderLayout.NORTH);
        sidebar.add(dayDetails, BorderLayout.CENTER);

        JPanel content = new JPanel(new BorderLayout());
        content.setBackground(CALENDAR_BACKGROUND);
        content.add(calendarContainer, BorderLayout.CENTER);
        content.add(sidebar, BorderLayout.EAST);
        return content;
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

            for (VlkDateRecord vlkDate : DatabaseManager.getAllVlkDates()) {
                LocalDate day = parseDate(vlkDate.getVlkDate());
                if (day == null) {
                    continue;
                }
                String responsible = vlkDate.getResponsible() == null ? "" : vlkDate.getResponsible();
                String event = vlkDate.getEventName() == null ? "" : vlkDate.getEventName();
                allEvents.add(new CalendarEvent(
                        day,
                        ASPECT_VLK,
                        "ВЛК: " + responsible,
                        event
                ));
            }

            SiteVisitRepository visitRepository = new SiteVisitRepository();
            Set<Integer> loadedVisitIds = new HashSet<>();
            for (int year : visitRepository.getAvailableYears()) {
                for (SiteVisitRecord visit : visitRepository.getVisitsForYear(year)) {
                    if (visit.status() == SiteVisitStatus.CANCELLED || !loadedVisitIds.add(visit.id())) {
                        continue;
                    }
                    addSiteVisitEvents(visit);
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

    private void addSiteVisitEvents(SiteVisitRecord visit) {
        if (visit.startAt() == null || visit.endAt() == null) {
            return;
        }
        LocalDate firstDay = visit.startAt().toLocalDate();
        LocalDate lastDay = visit.endAt().minusNanos(1).toLocalDate();
        if (lastDay.isBefore(firstDay)) {
            lastDay = firstDay;
        }

        String requestNumber = valueOrEmpty(visit.requestNumber());
        String title = requestNumber.isBlank() ? ASPECT_TESTS : ASPECT_TESTS + " · " + requestNumber;
        List<String> details = new ArrayList<>();
        if (!valueOrEmpty(visit.customerName()).isBlank()) {
            details.add(visit.customerName().trim());
        }
        if (!valueOrEmpty(visit.objectName()).isBlank()) {
            details.add(visit.objectName().trim());
        }
        if (!visit.allDay()) {
            details.add("Время: " + visit.startAt().format(TIME_LABEL) + "–" + visit.endAt().format(TIME_LABEL));
        }
        if (!visit.works().isEmpty()) {
            details.add("Работы: " + visit.works().stream()
                    .map(work -> work.name())
                    .filter(name -> name != null && !name.isBlank())
                    .collect(Collectors.joining(", ")));
        }
        String eventDetails = String.join("\n", details);

        for (LocalDate day = firstDay; !day.isAfter(lastDay); day = day.plusDays(1)) {
            allEvents.add(new CalendarEvent(day, ASPECT_TESTS, title, eventDetails));
        }
    }

    private static String valueOrEmpty(String value) {
        return value == null ? "" : value;
    }

    public void refreshEvents() {
        reloadEvents();
        renderCalendar();
        showSelectedDayDetails();
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
        JPanel root = new JPanel(new BorderLayout(0, 8));
        root.setBackground(CALENDAR_BACKGROUND);
        root.setBorder(new EmptyBorder(12, 14, 14, 14));
        root.add(createWeekHeader(), BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(6, 7, 1, 1));
        grid.setBackground(GRID_COLOR);
        grid.setBorder(BorderFactory.createLineBorder(GRID_COLOR));

        LocalDate firstDay = month.atDay(1);
        LocalDate start = firstDay.with(TemporalAdjusters.previousOrSame(DayOfWeek.MONDAY));
        for (int i = 0; i < 42; i++) {
            grid.add(createDayCell(start.plusDays(i), month, false));
        }

        root.add(grid, BorderLayout.CENTER);
        return root;
    }

    private JComponent createYearView(int year) {
        JPanel monthsGrid = new JPanel(new GridLayout(4, 3, 8, 8));
        monthsGrid.setBackground(CALENDAR_BACKGROUND);
        monthsGrid.setBorder(new EmptyBorder(12, 14, 14, 14));
        for (int m = 1; m <= 12; m++) {
            YearMonth ym = YearMonth.of(year, m);
            JPanel panel = new JPanel(new BorderLayout(0, 6));
            panel.setBackground(SURFACE_COLOR);
            JLabel label = new JLabel(capitalize(ym.format(DateTimeFormatter.ofPattern("LLLL", new Locale("ru")))), SwingConstants.LEFT);
            label.setFont(label.getFont().deriveFont(Font.BOLD, 13f));
            label.setBorder(new EmptyBorder(0, 2, 0, 0));
            panel.add(label, BorderLayout.NORTH);

            JPanel monthBody = new JPanel(new BorderLayout(0, 2));
            monthBody.setBackground(SURFACE_COLOR);
            monthBody.add(createWeekHeader(1, 1, 9f), BorderLayout.NORTH);

            JPanel mini = new JPanel(new GridLayout(6, 7, 1, 1));
            mini.setBackground(GRID_COLOR);
            LocalDate firstDay = ym.atDay(1);
            LocalDate start = firstDay.with(TemporalAdjusters.previousOrSame(DayOfWeek.MONDAY));
            for (int i = 0; i < 42; i++) {
                mini.add(createDayCell(start.plusDays(i), ym, true));
            }

            monthBody.add(mini, BorderLayout.CENTER);
            panel.add(monthBody, BorderLayout.CENTER);
            panel.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(GRID_COLOR),
                    new EmptyBorder(8, 8, 8, 8)
            ));
            monthsGrid.add(panel);
        }
        JScrollPane yearScroll = new JScrollPane(monthsGrid);
        yearScroll.setBorder(null);
        yearScroll.getViewport().setBackground(CALENDAR_BACKGROUND);
        yearScroll.getVerticalScrollBar().setUnitIncrement(32);
        yearScroll.getVerticalScrollBar().setBlockIncrement(128);
        return yearScroll;
    }

    private JComponent createWeekHeader() {
        return createWeekHeader(1, 0, 12f);
    }

    private JComponent createWeekHeader(int hGap, int vGap, float fontSize) {
        JPanel header = new JPanel(new GridLayout(1, 7, hGap, vGap));
        header.setOpaque(false);
        String[] days = {"Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"};
        for (int i = 0; i < days.length; i++) {
            String day = days[i];
            JLabel label = new JLabel(day, SwingConstants.CENTER);
            label.setFont(label.getFont().deriveFont(Font.BOLD, fontSize));
            label.setForeground(i >= 5 ? new Color(200, 79, 79) : MUTED_TEXT_COLOR);
            header.add(label);
        }
        return header;
    }

    private JComponent createDayCell(LocalDate date, YearMonth visibleMonth, boolean compact) {
        List<CalendarEvent> events = eventsForDate(date);

        JPanel cell = new JPanel(new BorderLayout(0, compact ? 0 : 3));
        cell.setBorder(new EmptyBorder(1, 1, 1, 1));
        cell.setBackground(SURFACE_COLOR);
        cell.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));

        boolean outsideMonth = !YearMonth.from(date).equals(visibleMonth);
        if (outsideMonth) {
            cell.setBackground(new Color(249, 250, 251));
        }
        if (date.equals(selectedDate)) {
            cell.setBackground(SELECTED_DAY_COLOR);
            cell.setBorder(BorderFactory.createLineBorder(PRIMARY_COLOR, compact ? 1 : 2));
        }

        JLabel dayNumber = new JLabel(String.valueOf(date.getDayOfMonth()));
        dayNumber.setHorizontalAlignment(SwingConstants.CENTER);
        dayNumber.setFont(dayNumber.getFont().deriveFont(compact ? 9f : 12f));
        dayNumber.setForeground(outsideMonth ? new Color(177, 182, 189) : Color.DARK_GRAY);
        dayNumber.setPreferredSize(new Dimension(compact ? 18 : 25, compact ? 17 : 25));
        if (date.equals(LocalDate.now())) {
            dayNumber.setOpaque(true);
            dayNumber.setBackground(PRIMARY_COLOR);
            dayNumber.setForeground(Color.WHITE);
            dayNumber.setBorder(new EmptyBorder(1, 4, 1, 4));
        }

        JPanel dayHeader = new JPanel(new FlowLayout(FlowLayout.RIGHT, compact ? 1 : 4, compact ? 0 : 3));
        dayHeader.setOpaque(false);
        dayHeader.add(dayNumber);
        cell.add(dayHeader, BorderLayout.NORTH);

        if (!compact) {
            JPanel eventsPanel = new JPanel();
            eventsPanel.setLayout(new BoxLayout(eventsPanel, BoxLayout.Y_AXIS));
            eventsPanel.setOpaque(false);
            eventsPanel.setBorder(new EmptyBorder(0, 4, 3, 4));

            int limit = 3;
            for (int i = 0; i < Math.min(limit, events.size()); i++) {
                CalendarEvent event = events.get(i);
                Color aspectColor = colorForAspect(event.aspect());
                JLabel eLabel = new JLabel(event.title());
                eLabel.setFont(eLabel.getFont().deriveFont(10f));
                eLabel.setForeground(darken(aspectColor));
                eLabel.setOpaque(true);
                eLabel.setBackground(tint(aspectColor));
                eLabel.setBorder(BorderFactory.createCompoundBorder(
                        BorderFactory.createMatteBorder(0, 4, 0, 0, aspectColor),
                        new EmptyBorder(2, 5, 2, 4)
                ));
                eLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
                eLabel.setMaximumSize(new Dimension(Integer.MAX_VALUE, 22));
                eLabel.setToolTipText(event.title());
                eventsPanel.add(eLabel);
                eventsPanel.add(Box.createVerticalStrut(2));
            }
            if (events.size() > limit) {
                JLabel more = new JLabel("Ещё " + (events.size() - limit));
                more.setFont(more.getFont().deriveFont(Font.ITALIC, 10f));
                more.setForeground(PRIMARY_COLOR);
                more.setBorder(new EmptyBorder(1, 5, 0, 0));
                eventsPanel.add(more);
            }
            cell.add(eventsPanel, BorderLayout.CENTER);
        } else if (!events.isEmpty()) {
            JPanel marker = createEventMarkerPanel(events, compact);
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
        detailsEventsPanel.removeAll();
        List<CalendarEvent> events = eventsForDate(selectedDate);
        if (events.isEmpty()) {
            JLabel emptyLabel = new JLabel("<html>На выбранную дату<br>событий нет</html>");
            emptyLabel.setForeground(MUTED_TEXT_COLOR);
            emptyLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
            detailsEventsPanel.add(emptyLabel);
            detailsEventsPanel.revalidate();
            detailsEventsPanel.repaint();
            return;
        }

        for (CalendarEvent event : events) {
            detailsEventsPanel.add(createEventDetailsCard(event));
            detailsEventsPanel.add(Box.createVerticalStrut(8));
        }
        detailsEventsPanel.add(Box.createVerticalGlue());
        detailsEventsPanel.revalidate();
        detailsEventsPanel.repaint();
    }

    private JComponent createEventDetailsCard(CalendarEvent event) {
        Color aspectColor = colorForAspect(event.aspect());
        Color background = tint(aspectColor);

        JLabel aspectLabel = new JLabel(event.aspect());
        aspectLabel.setForeground(darken(aspectColor));
        aspectLabel.setFont(aspectLabel.getFont().deriveFont(Font.BOLD, 11f));

        JTextArea title = createDetailsText(event.title(), background, Font.BOLD);
        JTextArea details = createDetailsText(event.details(), background, Font.PLAIN);

        JPanel card = new JPanel();
        card.setLayout(new BoxLayout(card, BoxLayout.Y_AXIS));
        card.setBackground(background);
        card.setAlignmentX(Component.LEFT_ALIGNMENT);
        card.setMaximumSize(new Dimension(Integer.MAX_VALUE, 124));
        card.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createMatteBorder(0, 4, 0, 0, aspectColor),
                new EmptyBorder(8, 10, 8, 10)
        ));
        aspectLabel.setAlignmentX(Component.LEFT_ALIGNMENT);
        title.setAlignmentX(Component.LEFT_ALIGNMENT);
        details.setAlignmentX(Component.LEFT_ALIGNMENT);
        card.add(aspectLabel);
        card.add(Box.createVerticalStrut(4));
        card.add(title);
        if (event.details() != null && !event.details().isBlank()) {
            card.add(Box.createVerticalStrut(3));
            card.add(details);
        }
        return card;
    }

    private JTextArea createDetailsText(String text, Color background, int fontStyle) {
        JTextArea area = new JTextArea(text == null ? "" : text);
        area.setEditable(false);
        area.setFocusable(false);
        area.setLineWrap(true);
        area.setWrapStyleWord(true);
        area.setOpaque(true);
        area.setBackground(background);
        area.setBorder(null);
        area.setFont(area.getFont().deriveFont(fontStyle, fontStyle == Font.BOLD ? 12f : 11f));
        area.setRows(1);
        area.setColumns(20);
        return area;
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

    private JPanel createEventMarkerPanel(List<CalendarEvent> events, boolean compact) {
        JPanel marker = new JPanel(new GridLayout(1, 0, 1, 0));
        marker.setOpaque(false);

        int limit = compact ? 3 : 4;
        Set<String> uniqueAspects = events.stream()
                .map(CalendarEvent::aspect)
                .collect(Collectors.toCollection(LinkedHashSet::new));

        int count = 0;
        for (String aspect : uniqueAspects) {
            if (count++ >= limit) {
                break;
            }
            JPanel strip = new JPanel();
            strip.setOpaque(true);
            strip.setBackground(colorForAspect(aspect));
            strip.setPreferredSize(new Dimension(8, compact ? 4 : 6));
            marker.add(strip);
        }
        return marker;
    }

    private Color colorForAspect(String aspect) {
        return ASPECT_COLORS.getOrDefault(aspect, new Color(120, 144, 156));
    }

    private static Color tint(Color color) {
        return new Color(
                (color.getRed() + 255 * 6) / 7,
                (color.getGreen() + 255 * 6) / 7,
                (color.getBlue() + 255 * 6) / 7
        );
    }

    private static Color darken(Color color) {
        return new Color(
                Math.max(0, color.getRed() * 2 / 3),
                Math.max(0, color.getGreen() * 2 / 3),
                Math.max(0, color.getBlue() * 2 / 3)
        );
    }

    private static LocalDate parseDate(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        String trimmed = value.trim();

        if (trimmed.length() >= 10) {
            try {
                return LocalDate.parse(trimmed.substring(0, 10));
            } catch (DateTimeParseException ignored) {
                // Пробуем альтернативные форматы ниже.
            }
        }

        try {
            return LocalDate.parse(trimmed);
        } catch (DateTimeParseException ignored) {
            try {
                return LocalDate.parse(trimmed, DOTTED_DATE_FORMAT);
            } catch (DateTimeParseException ignoredAgain) {
                try {
                    return LocalDate.parse(trimmed, DASHED_DATE_FORMAT);
                } catch (DateTimeParseException ignoredThird) {
                    return null;
                }
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
