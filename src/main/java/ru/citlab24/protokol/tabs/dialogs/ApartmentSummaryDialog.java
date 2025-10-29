package ru.citlab24.protokol.tabs.dialogs;

import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Section;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Space;
import ru.citlab24.protokol.tabs.models.Room;

import javax.swing.*;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellRenderer;
import java.awt.*;
import java.util.List;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * Сводка по квартирам: теперь по СЕКЦИЯМ.
 * Для каждой секции — собственная вкладка со сводкой и матрицей.
 */
public final class ApartmentSummaryDialog extends JDialog {

    public static void show(Component parent, Building building) {
        Window owner = parent == null ? null : SwingUtilities.getWindowAncestor(parent);
        ApartmentSummaryDialog d = new ApartmentSummaryDialog(owner, building);
        d.setVisible(true);
    }

    private final Building building;

    private ApartmentSummaryDialog(Window owner, Building building) {
        super(owner, "Сводка по квартирам", ModalityType.APPLICATION_MODAL);
        this.building = (building != null ? building : new Building());
        setDefaultCloseOperation(DISPOSE_ON_CLOSE);
        setLayout(new BorderLayout(8, 8));

        // ===== Вкладки по секциям =====
        JTabbedPane tabs = new JTabbedPane();

        List<Section> sections = safeSections(this.building);
        if (sections.isEmpty()) {
            // кейс "без секций": показываем одну вкладку "Здание"
            tabs.addTab("Здание", buildSectionPanel(-1, "Здание"));
        } else {
            // сортируем по position/имени
            sections.sort(Comparator.comparingInt(s -> Optional.ofNullable(s.getPosition()).orElse(0)));
            for (int i = 0; i < sections.size(); i++) {
                Section s = sections.get(i);
                String title = (s.getName() == null || s.getName().isBlank())
                        ? ("Секция " + (i + 1))
                        : s.getName();
                tabs.addTab(title, buildSectionPanel(i, title));
            }
        }

        add(tabs, BorderLayout.CENTER);

        JButton close = new JButton("Закрыть");
        close.addActionListener(e -> dispose());
        JPanel btns = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        btns.add(close);
        add(btns, BorderLayout.SOUTH);

        setMinimumSize(new Dimension(1000, 650));
        setLocationRelativeTo(owner);
    }

    /** Панель для одной секции (index = -1 означает "все этажи без фильтра по секции"). */
    private JComponent buildSectionPanel(int sectionIndex, String sectionTitle) {
        // 1) собираем квартиры этой секции по номеру этажа
        Map<Integer, List<Space>> aptsByFloorNum = collectApartmentsByFloorNumber(this.building, sectionIndex);

        List<Integer> floorNums = new ArrayList<>(aptsByFloorNum.keySet());
        Collections.sort(floorNums);

        int totalApts = aptsByFloorNum.values().stream().mapToInt(List::size).sum();

        // 2) поэтажная сводка
        List<FloorSummary> summaries = new ArrayList<>();
        for (Integer fn : floorNums) {
            List<Space> apts = aptsByFloorNum.getOrDefault(fn, List.of());
            int rooms = 0;
            int baths = 0;
            for (Space s : apts) {
                List<Room> rr = safeRooms(s);
                rooms += rr.size();
                baths += (int) rr.stream().filter(ApartmentSummaryDialog::isBathroom).count();
            }
            summaries.add(new FloorSummary(fn, apts.size(), rooms, baths));
        }

        // 3) верх: общее количество + таблица-сводка
        JLabel totalLabel = new JLabel("Общее количество квартир: " + totalApts + "   —   " + sectionTitle);
        totalLabel.setFont(totalLabel.getFont().deriveFont(Font.BOLD, 14f));

        JTable tSummary = new JTable(new FloorSummaryModel(summaries));
        tSummary.setFillsViewportHeight(true);
        tSummary.setRowHeight(28);
        JScrollPane spSummary = new JScrollPane(tSummary);
        spSummary.setBorder(BorderFactory.createTitledBorder("Сводка по этажам"));

        // заголовок + кнопка "Развернуть матрицу"
        JButton toggleMatrixBtn = new JButton("Развернуть матрицу");
        final double[] lastDivider = {0.45};      // запомним стандартное положение
        final boolean[] expandedOnlyMatrix = {false};

        JPanel header = new JPanel(new BorderLayout(8, 0));
        header.add(totalLabel, BorderLayout.WEST);
        header.add(toggleMatrixBtn, BorderLayout.EAST);

        JPanel top = new JPanel(new BorderLayout(6, 6));
        top.add(header, BorderLayout.NORTH);
        top.add(spSummary, BorderLayout.CENTER);


        // 4) низ: матрица этажи × квартиры
        MatrixModel matrixModel = new MatrixModel(floorNums, aptsByFloorNum);
        JTable tMatrix = new JTable(matrixModel);
        tMatrix.setDefaultRenderer(Object.class, new SingleLineCellRenderer()); // ← новый рендерер (ниже)
        tMatrix.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);

// фиксированная высота = одна строка текста
        tMatrix.setRowHeight(computeSingleLineRowHeight(tMatrix));

// ширины колонок
        tMatrix.getColumnModel().getColumn(0).setPreferredWidth(80);
        for (int i = 1; i < tMatrix.getColumnModel().getColumnCount(); i++) {
            tMatrix.getColumnModel().getColumn(i).setPreferredWidth(200);
        }

        JScrollPane spMatrix = new JScrollPane(tMatrix);
        spMatrix.setBorder(BorderFactory.createTitledBorder(
                "Матрица: этажи × квартиры (ячейка — комнаты квартиры)"));

        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT, top, spMatrix);
        split.setResizeWeight(0.45);
        split.setContinuousLayout(true);

// обработчик "Развернуть матрицу" ↔ "Показать сводку"
        toggleMatrixBtn.addActionListener(ev -> {
            if (!expandedOnlyMatrix[0]) {
                lastDivider[0] = Math.max(0.1, split.getDividerLocation() / Math.max(1.0, split.getHeight()));
                split.setDividerLocation(0.0);                 // всю высоту отдаём матрице
                toggleMatrixBtn.setText("Показать сводку");
                expandedOnlyMatrix[0] = true;
            } else {
                split.setDividerLocation(lastDivider[0]);      // вернуть как было (ok и в процентах, и в пикселях)
                toggleMatrixBtn.setText("Развернуть матрицу");
                expandedOnlyMatrix[0] = false;
            }
        });

        JPanel page = new JPanel(new BorderLayout());
        page.add(split, BorderLayout.CENTER);
        return page;

    }

    // ===== модели =====
    private static final class FloorSummary {
        final int floorNum, apartments, rooms, bathrooms;
        FloorSummary(int fn, int a, int r, int b) { floorNum = fn; apartments = a; rooms = r; bathrooms = b; }
    }
    private static final class FloorSummaryModel extends AbstractTableModel {
        private final String[] cols = {"Этаж", "Квартир", "Комнат", "Санузлов"};
        private final List<FloorSummary> rows;
        FloorSummaryModel(List<FloorSummary> rows) { this.rows = rows; }
        @Override public int getRowCount() { return rows.size(); }
        @Override public int getColumnCount() { return cols.length; }
        @Override public String getColumnName(int c) { return cols[c]; }
        @Override public Class<?> getColumnClass(int c) { return Integer.class; }
        @Override public Object getValueAt(int r, int c) {
            FloorSummary x = rows.get(r);
            return switch (c) { case 0 -> x.floorNum; case 1 -> x.apartments; case 2 -> x.rooms; case 3 -> x.bathrooms; default -> ""; };
        }
    }
    private static final class MatrixModel extends AbstractTableModel {
        private final List<Integer> floorNums;
        private final Map<Integer, List<Space>> aptsByFloor;
        private final int maxApts;
        MatrixModel(List<Integer> floorNums, Map<Integer, List<Space>> aptsByFloor) {
            this.floorNums = floorNums;
            this.aptsByFloor = aptsByFloor;
            this.maxApts = aptsByFloor.values().stream().mapToInt(List::size).max().orElse(0);
        }
        @Override public int getRowCount() { return floorNums.size(); }
        @Override public int getColumnCount() { return 1 + Math.max(1, maxApts); }
        @Override public String getColumnName(int c) { return (c == 0) ? "Этаж" : ("Кв. " + c); }
        @Override public Class<?> getColumnClass(int c) { return Object.class; }
        @Override public Object getValueAt(int r, int c) {
            int fnum = floorNums.get(r);
            if (c == 0) return fnum;
            List<Space> apts = aptsByFloor.getOrDefault(fnum, List.of());
            int idx = c - 1;
            if (idx < 0 || idx >= apts.size()) return "";
            Space apt = apts.get(idx);
            List<Room> rooms = safeRooms(apt);
            if (rooms.isEmpty()) return "—";
            return rooms.stream().map(x -> abbreviate(x.getName())).collect(Collectors.joining(", "));
        }
    }

    // ===== сбор данных =====
    private static Map<Integer, List<Space>> collectApartmentsByFloorNumber(Building b, int sectionIndex) {
        Map<Integer, List<Space>> map = new HashMap<>();
        if (b == null) return map;
        List<Floor> floors = (b.getFloors() != null) ? b.getFloors() : List.of();

        for (Floor f : floors) {
            if (sectionIndex >= 0 && f.getSectionIndex() != sectionIndex) continue; // фильтр по секции
            int floorNum = parseFloorNumber(f);
            if (floorNum == Integer.MIN_VALUE) continue;

            List<Space> spaces = (f.getSpaces() != null) ? f.getSpaces() : List.of();
            for (Space s : spaces) {
                if (!isApartment(s)) continue;
                map.computeIfAbsent(floorNum, k -> new ArrayList<>()).add(s);
            }
        }

        // стабильный порядок квартир на этаже
        for (List<Space> list : map.values()) {
            list.sort((a, c) -> {
                int byId = safeId(a).compareToIgnoreCase(safeId(c));
                if (byId != 0) return byId;
                int rc = Integer.compare(safeRooms(c).size(), safeRooms(a).size());
                if (rc != 0) return rc;
                return 0;
            });
        }
        return map;
    }

    private static List<Section> safeSections(Building b) {
        try { List<Section> s = b.getSections(); return (s == null) ? List.of() : s; }
        catch (Throwable t) { return List.of(); }
    }
    private static String safeId(Space s) {
        String id = (s != null && s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        return id.isEmpty() ? "~" : id;
    }
    private static int parseFloorNumber(Floor f) {
        try { Integer n = extractFirstInt(f.getNumber()); if (n != null) return n; } catch (Throwable ignore) {}
        try { Integer n = extractFirstInt(f.getName());   if (n != null) return n; } catch (Throwable ignore) {}
        return Integer.MIN_VALUE;
    }
    private static Integer extractFirstInt(String s) {
        if (s == null) return null;
        Matcher m = Pattern.compile("(-?\\d+)").matcher(s);
        return m.find() ? Integer.parseInt(m.group(1)) : null;
    }
    private static List<Room> safeRooms(Space s) {
        try { List<Room> r = s.getRooms(); return (r == null) ? List.of() : r; }
        catch (Throwable t) { return List.of(); }
    }
    private static boolean isApartment(Space s) {
        if (s == null) return false;
        try {
            Space.SpaceType t = s.getType();
            if (t != null) {
                String tn = t.name();
                if ("APARTMENT".equalsIgnoreCase(tn) ||
                        "FLAT".equalsIgnoreCase(tn) ||
                        "RESIDENTIAL".equalsIgnoreCase(tn) ||
                        "LIVING".equalsIgnoreCase(tn)) return true;
            }
        } catch (Throwable ignore) {}
        String id = (s.getIdentifier() == null) ? "" : s.getIdentifier().toLowerCase(Locale.ROOT);
        return id.contains("кв") || id.contains("квартир") || id.contains("apartment") || id.startsWith("apt");
    }
    private static boolean isBathroom(Room r) {
        String n = (r == null || r.getName() == null) ? "" : r.getName().toLowerCase(Locale.ROOT);
        return n.contains("сануз") || n.contains("с/у") || n.contains("ванн") || n.contains("душ") || n.contains("туал") || n.contains("уборн");
    }

    // ===== сокращения =====
    private static final LinkedHashMap<String, String> ABBR = new LinkedHashMap<>();
    static {
        ABBR.put("жила", "жк");   // ← явный запрос: «жилая комната» → «жк»
        ABBR.put("спальн", "СП");
        ABBR.put("гости", "ГОС");
        ABBR.put("детск", "ДЕТ");
        ABBR.put("корид", "КОР");
        ABBR.put("прихож","ПРХ");
        ABBR.put("сануз", "С/У");
        ABBR.put("с/у",   "С/У");
        ABBR.put("ванн",  "ВН");
        ABBR.put("совмещ",  "СС/У");
        ABBR.put("туал",  "ТУ");
        ABBR.put("гардер","ГРД");
        ABBR.put("балкон","БЛК");
        ABBR.put("кабин", "КАБ");
        ABBR.put("клад",  "КЛАД");
    }
    private static String abbreviate(String name) {
        if (name == null || name.isBlank()) return "";
        String low = name.toLowerCase(Locale.ROOT);
        for (var e : ABBR.entrySet()) if (low.contains(e.getKey())) return e.getValue();
        return name;
    }
    // === Авто-упаковка высоты строк таблицы под содержимое ===
    private static void packRowHeights(JTable table, int margin) {
        for (int row = 0; row < table.getRowCount(); row++) {
            int h = computePreferredRowHeight(table, row, margin);
            if (h > 0 && table.getRowHeight(row) != h) {
                table.setRowHeight(row, h);
            }
        }
    }
    private static int computePreferredRowHeight(JTable table, int row, int margin) {
        int max = 0;
        for (int col = 0; col < table.getColumnCount(); col++) {
            TableCellRenderer r = table.getCellRenderer(row, col);
            Component c = table.prepareRenderer(r, row, col);
            int h = c.getPreferredSize().height + margin;
            max = Math.max(max, h);
        }
        return Math.max(max, 1);
    }
    /** Рендерер одной строки: текст не переносится, лишнее обрезается с «…», полный текст — в tooltip. */
    private static final class SingleLineCellRenderer extends JLabel implements TableCellRenderer {
        private static final String ELLIPSIS = "…";

        SingleLineCellRenderer() {
            setOpaque(true);
            setBorder(BorderFactory.createEmptyBorder(4, 6, 4, 6));
            setVerticalAlignment(CENTER);
            setHorizontalAlignment(LEFT);
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value,
                                                       boolean isSelected, boolean hasFocus,
                                                       int row, int column) {
            String full = (value == null) ? "" : String.valueOf(value);

            // обрезаем до одной строки с "…"
            int colWidth = table.getColumnModel().getColumn(column).getWidth();
            int available = Math.max(0, colWidth - getInsets().left - getInsets().right - 6);
            String shown = clipText(full, available, table.getFontMetrics(getFont()));
            setText(shown);
            setToolTipText(full.isEmpty() ? null : full);

            // --- ПОДСВЕТКА: отличается от ячейки НА СЛЕДУЮЩЕМ ЭТАЖЕ в той же квартире ---
            boolean highlight = false;
            if (column > 0 && row < table.getRowCount() - 1) { // колонки квартир начинаются с 1
                Object nextObj = table.getValueAt(row + 1, column);
                String nextFull = (nextObj == null) ? "" : String.valueOf(nextObj);
                // считаем «одинаково», если совпадает без учёта регистра и лишних пробелов вокруг запятых
                highlight = !equalsNormalized(full, nextFull);
            }

            if (isSelected) {
                setBackground(table.getSelectionBackground());
                setForeground(table.getSelectionForeground());
            } else {
                setBackground(highlight ? HILIGHT_BG : table.getBackground());
                setForeground(table.getForeground());
            }
            return this;
        }


        /** Обрезает строку под пиксельную ширину, добавляя «…» при необходимости (быстрый бинпоиск). */
        private static String clipText(String s, int maxW, FontMetrics fm) {
            if (s == null || s.isEmpty()) return "";
            if (fm.stringWidth(s) <= maxW) return s;

            int lo = 0, hi = s.length();
            while (lo < hi) {
                int mid = (lo + hi) >>> 1;
                String cand = s.substring(0, Math.max(0, mid)) + ELLIPSIS;
                if (fm.stringWidth(cand) <= maxW) lo = mid + 1; else hi = mid;
            }
            int cut = Math.max(0, lo - 1);
            if (cut <= 0) return ELLIPSIS;                     // даже 1 символ не помещается
            return s.substring(0, cut) + ELLIPSIS;
        }
    }
    /** Высота строки таблицы под одну строку текста с небольшим внутренним отступом. */
    private static int computeSingleLineRowHeight(JTable t) {
        FontMetrics fm = t.getFontMetrics(t.getFont());
        int text = fm.getAscent() + fm.getDescent();
        return Math.max(20, text + 8);  // 8 — запас под паддинги/бордер
    }
    // нежно-голубой фон для подсветки отличий
    private static final Color HILIGHT_BG = new Color(225, 245, 254);

    /** Нормализованное сравнение содержимого ячеек матрицы. */
    private static boolean equalsNormalized(String a, String b) {
        return normalizeCell(a).equals(normalizeCell(b));
    }
    private static String normalizeCell(String s) {
        if (s == null) return "";
        // нижний регистр, обрезка и нормализация ", " → ","
        String t = s.toLowerCase(Locale.ROOT).trim();
        t = t.replaceAll("\\s*,\\s*", ",");   // "жк, ВН, кухня" -> "жк,вн,кухня"
        t = t.replaceAll("\\s+", " ");       // лишние пробелы
        return t;
    }

}
