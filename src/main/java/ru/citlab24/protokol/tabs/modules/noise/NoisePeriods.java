package ru.citlab24.protokol.tabs.modules.noise;

import javax.swing.*;
import java.awt.*;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.EnumMap;
import java.util.Map;

/** Период измерений: дата + время «с/до». */
final class NoisePeriod {
    LocalDate date;
    LocalTime from;
    LocalTime to;

    NoisePeriod() { }
    NoisePeriod(LocalDate d, LocalTime f, LocalTime t) { this.date=d; this.from=f; this.to=t; }

    /** Текст для Excel: «Дата, время проведения измерений dd.MM.yyyy c HH:mm до HH:mm». */
    String toExcelLine() {
        DateTimeFormatter df = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        DateTimeFormatter tf = DateTimeFormatter.ofPattern("HH:mm");
        String d = (date != null)? date.format(df) : "__.__.____";
        String f = (from != null)? from.format(tf) : "__:__";
        String t = (to   != null)? to.format(tf)   : "__:__";
        return "Дата, время проведения измерений " + d + " c " + f + " до " + t;
    }
}

/** Диалог ввода периодов для всех нужных испытаний: лифт/ИТО/авто/площадка. */
final class NoisePeriodsDialog extends JDialog {
    private final Map<NoiseTestKind, NoisePeriod> result = new EnumMap<>(NoiseTestKind.class);

    // Лифт
    private final PeriodPanel pLiftDay      = new PeriodPanel("Лифт — день",         LocalTime.of(8, 0),  LocalTime.of(21, 0));
    private final PeriodPanel pLiftNight    = new PeriodPanel("Лифт — ночь",         LocalTime.of(22, 0), LocalTime.of(6,  0));

    // ИТО
    private final PeriodPanel pItoNonres    = new PeriodPanel("ИТО — нежилые",       LocalTime.of(8, 0),  LocalTime.of(21, 0));
    private final PeriodPanel pItoResDay    = new PeriodPanel("ИТО — жилые день",    LocalTime.of(8, 0),  LocalTime.of(21, 0));  // новое имя
    private final PeriodPanel pItoResNight  = new PeriodPanel("ИТО — жилые ночь",    LocalTime.of(22, 0), LocalTime.of(6,  0));  // новый пункт

    // ЗУМ
    private final PeriodPanel pZumDay       = new PeriodPanel("ЗУМ — день",          LocalTime.of(8, 0),  LocalTime.of(21, 0));

    // Авто
    private final PeriodPanel pAutoDay      = new PeriodPanel("Авто — день",         LocalTime.of(8, 0),  LocalTime.of(21, 0));
    private final PeriodPanel pAutoNight    = new PeriodPanel("Авто — ночь",         LocalTime.of(22, 0), LocalTime.of(6,  0));

    // Площадка (улица)
    private final PeriodPanel pSite         = new PeriodPanel("Площадка (улица)",    LocalTime.of(8, 0),  LocalTime.of(21, 0));

    NoisePeriodsDialog(Window owner,
                       Map<NoiseTestKind, NoisePeriod> initial,
                       Map<NoiseTestKind, Integer> points) {
        super(owner, "Периоды измерений (шумы)", ModalityType.APPLICATION_MODAL);
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);

        JPanel center = new JPanel();
        center.setLayout(new BoxLayout(center, BoxLayout.Y_AXIS));
        center.setBorder(BorderFactory.createEmptyBorder(8, 8, 8, 8));

        center.add(makeGroup("Лифт", pLiftDay, pLiftNight));
        center.add(Box.createVerticalStrut(8));

        center.add(makeGroup("ИТО", pItoNonres, pItoResDay, pItoResNight)); // было pItoRes
        center.add(Box.createVerticalStrut(8));

        center.add(makeGroup("ЗУМ", pZumDay));
        center.add(Box.createVerticalStrut(8));

        center.add(makeGroup("Авто", pAutoDay, pAutoNight));
        center.add(Box.createVerticalStrut(8));

        center.add(makeGroup("Площадка (улица)", pSite));

        JButton ok = new JButton("OK");
        JButton cancel = new JButton("Отмена");
        JPanel south = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        south.add(cancel);
        south.add(ok);

        ok.addActionListener(e -> {
            result.clear();
            result.put(NoiseTestKind.LIFT_DAY,      pLiftDay.getPeriod());
            result.put(NoiseTestKind.LIFT_NIGHT,    pLiftNight.getPeriod());

            result.put(NoiseTestKind.ITO_NONRES,    pItoNonres.getPeriod());
            result.put(NoiseTestKind.ITO_RES_DAY,   pItoResDay.getPeriod());    // новое
            result.put(NoiseTestKind.ITO_RES_NIGHT, pItoResNight.getPeriod());  // новое

            result.put(NoiseTestKind.ZUM_DAY,       pZumDay.getPeriod());

            result.put(NoiseTestKind.AUTO_DAY,      pAutoDay.getPeriod());
            result.put(NoiseTestKind.AUTO_NIGHT,    pAutoNight.getPeriod());

            result.put(NoiseTestKind.SITE,          pSite.getPeriod());
            dispose();
        });
        cancel.addActionListener(e -> {
            if (initial != null) result.putAll(initial);
            dispose();
        });

        // Проставим начальные значения, если были
        if (initial != null) {
            setIfPresent(initial, NoiseTestKind.LIFT_DAY,      pLiftDay);
            setIfPresent(initial, NoiseTestKind.LIFT_NIGHT,    pLiftNight);
            setIfPresent(initial, NoiseTestKind.ITO_NONRES,    pItoNonres);
            setIfPresent(initial, NoiseTestKind.ITO_RES_DAY,   pItoResDay);     // новое
            setIfPresent(initial, NoiseTestKind.ITO_RES_NIGHT, pItoResNight);   // новое
            setIfPresent(initial, NoiseTestKind.ZUM_DAY,       pZumDay);
            setIfPresent(initial, NoiseTestKind.AUTO_DAY,      pAutoDay);
            setIfPresent(initial, NoiseTestKind.AUTO_NIGHT,    pAutoNight);
            setIfPresent(initial, NoiseTestKind.SITE,          pSite);
        }
        if (points != null) {
            setPointsIfPresent(points, NoiseTestKind.LIFT_DAY, pLiftDay);
            setPointsIfPresent(points, NoiseTestKind.LIFT_NIGHT, pLiftNight);
            setPointsIfPresent(points, NoiseTestKind.ITO_NONRES, pItoNonres);
            setPointsIfPresent(points, NoiseTestKind.ITO_RES_DAY, pItoResDay);
            setPointsIfPresent(points, NoiseTestKind.ITO_RES_NIGHT, pItoResNight);
            setPointsIfPresent(points, NoiseTestKind.ZUM_DAY, pZumDay);
            setPointsIfPresent(points, NoiseTestKind.AUTO_DAY, pAutoDay);
            setPointsIfPresent(points, NoiseTestKind.AUTO_NIGHT, pAutoNight);
            setPointsIfPresent(points, NoiseTestKind.SITE, pSite);
        }

        getContentPane().setLayout(new BorderLayout(8, 8));
        getContentPane().add(center, BorderLayout.CENTER);
        getContentPane().add(south,  BorderLayout.SOUTH);

        setPreferredSize(new Dimension(560, 800));
        pack();
        setLocationRelativeTo(owner);
    }

    private static JPanel makeGroup(String title, JComponent... items) {
        JPanel group = new JPanel(new GridBagLayout());
        group.setBorder(BorderFactory.createTitledBorder(title));
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(4, 6, 4, 6);
        c.gridx = 0; c.gridy = 0; c.anchor = GridBagConstraints.WEST; c.fill = GridBagConstraints.HORIZONTAL;
        for (JComponent it : items) {
            group.add(it, c);
            c.gridy++;
        }
        return group;
    }

    private static void setIfPresent(Map<NoiseTestKind, NoisePeriod> m, NoiseTestKind k, PeriodPanel p) {
        NoisePeriod np = m.get(k);
        if (np != null) p.setPeriod(np);
    }

    private static void setPointsIfPresent(Map<NoiseTestKind, Integer> m, NoiseTestKind k, PeriodPanel p) {
        if (m == null || k == null || p == null) return;
        Integer points = m.get(k);
        p.setPointsCount((points != null) ? points : 0);
    }

    Map<NoiseTestKind, NoisePeriod> getResult() { return result; }

    /* ====== панель одной строки: дата + с/до ====== */
    private static final class PeriodPanel extends JPanel {
        private final JSpinner dDate = new JSpinner(new SpinnerDateModel());
        private final JSpinner tFrom = new JSpinner(new SpinnerDateModel());
        private final JSpinner tTo   = new JSpinner(new SpinnerDateModel());
        private final JLabel pointsLabel = new JLabel();

        PeriodPanel(String title, LocalTime defFrom, LocalTime defTo) {
            super(new GridBagLayout());
            setBorder(BorderFactory.createTitledBorder(title));

            JSpinner.DateEditor edDate = new JSpinner.DateEditor(dDate, "dd.MM.yyyy");
            dDate.setEditor(edDate);
            edDate.getFormat().setLenient(false);
            edDate.getTextField().setColumns(10);

            JSpinner.DateEditor edFrom = new JSpinner.DateEditor(tFrom, "HH:mm");
            tFrom.setEditor(edFrom);
            edFrom.getFormat().setLenient(false);
            edFrom.getTextField().setColumns(5);

            JSpinner.DateEditor edTo = new JSpinner.DateEditor(tTo, "HH:mm");
            tTo.setEditor(edTo);
            edTo.getFormat().setLenient(false);
            edTo.getTextField().setColumns(5);

            ZonedDateTime now = ZonedDateTime.now();
            dDate.setValue(java.util.Date.from(now.toInstant()));
            tFrom.setValue(java.util.Date.from(now.with(defFrom).toInstant()));
            tTo.setValue(java.util.Date.from(now.with(defTo).toInstant()));

            GridBagConstraints c = new GridBagConstraints();
            c.insets = new Insets(4, 6, 4, 6);
            c.gridy = 0;

            c.gridx = 0; add(new JLabel("Дата:"), c);
            c.gridx = 1; add(dDate, c);
            c.gridx = 2; add(new JLabel("с:"), c);
            c.gridx = 3; add(tFrom, c);
            c.gridx = 4; add(new JLabel("до:"), c);
            c.gridx = 5; add(tTo, c);

            setPointsCount(0);
            c.gridx = 6; c.fill = GridBagConstraints.NONE; c.weightx = 0;
            add(pointsLabel, c);

            c.weightx = 1; c.gridx = 7; c.fill = GridBagConstraints.HORIZONTAL;
            add(Box.createHorizontalStrut(1), c);
        }

        NoisePeriod getPeriod() {
            ZoneId z = ZoneId.systemDefault();
            java.util.Date dd = (java.util.Date) dDate.getValue();
            java.util.Date ff = (java.util.Date) tFrom.getValue();
            java.util.Date tt = (java.util.Date) tTo.getValue();

            LocalDate d = ZonedDateTime.ofInstant(dd.toInstant(), z).toLocalDate();
            LocalTime f = ZonedDateTime.ofInstant(ff.toInstant(), z).toLocalTime().withSecond(0).withNano(0);
            LocalTime t = ZonedDateTime.ofInstant(tt.toInstant(), z).toLocalTime().withSecond(0).withNano(0);
            return new NoisePeriod(d, f, t);
        }

        void setPeriod(NoisePeriod p) {
            if (p == null) return;
            ZoneId z = ZoneId.systemDefault();
            LocalDate d = (p.date != null) ? p.date : LocalDate.now();
            LocalTime f = (p.from != null) ? p.from : LocalTime.of(8, 0);
            LocalTime t = (p.to   != null) ? p.to   : LocalTime.of(21, 0);
            dDate.setValue(java.util.Date.from(d.atStartOfDay(z).toInstant()));
            tFrom.setValue(java.util.Date.from(d.atTime(f).atZone(z).toInstant()));
            tTo.setValue(java.util.Date.from(d.atTime(t).atZone(z).toInstant()));
        }

        void setPointsCount(int points) {
            pointsLabel.setText("Точек: " + points);
        }
    }
}
