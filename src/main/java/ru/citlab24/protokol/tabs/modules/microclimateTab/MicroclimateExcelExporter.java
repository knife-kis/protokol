package ru.citlab24.protokol.tabs.modules.microclimateTab;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public final class MicroclimateExcelExporter {

    private MicroclimateExcelExporter() {}

    public enum TemperatureMode {
        COLD("Холодно", 0.0),
        NORMAL("Нормально", 0.4),
        WARM("Тепло", 1.0);

        private final String label;
        private final double offset;

        TemperatureMode(String label, double offset) {
            this.label = label;
            this.offset = offset;
        }

        public String getLabel() {
            return label;
        }

        public double getOffset() {
            return offset;
        }
    }

    /* ===== ПУБЛИЧНЫЕ API ===== */

    /** append: строит «большой» лист точь-в-точь как при экспорте из вкладки. */
    public static void appendToWorkbook(Building building, int sectionIndex, Workbook wb, TemperatureMode mode) {
        if (building == null || wb == null) return;
        buildMicroclimateSheets(wb, building, sectionIndex, mode);
    }

    /** export: обёртка — создаёт книгу, вызывает builder и предлагает сохранить. */
    public static void export(Building building, int sectionIndex, TemperatureMode mode, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        try (Workbook wb = new XSSFWorkbook()) {
            buildMicroclimateSheets(wb, building, sectionIndex, mode);

            // ===== сохранение =====
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Микроклимат.xlsx"));
            if (chooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = chooser.getSelectedFile();
                try (FileOutputStream out = new FileOutputStream(file)) {
                    wb.write(out);
                }
                JOptionPane.showMessageDialog(parent,
                        "Файл сохранён:\n" + file.getAbsolutePath(),
                        "Экспорт завершён", JOptionPane.INFORMATION_MESSAGE);
            }
        } catch (Exception ex) {
            ex.printStackTrace();
            JOptionPane.showMessageDialog(parent, "Ошибка экспорта: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }

    /* ===== ВЕСЬ СТРОИТЕЛЬ ЛИСТА — как во вкладке ===== */

    /** Весь «большой» шаблон (A..V, объединения, стили, допуски и т. п.). */
    private static void buildMicroclimateSheets(Workbook wb, Building building, int sectionIndex, TemperatureMode mode) {
        Styles S = new Styles(wb);
        Sheet sh = wb.createSheet(uniqueName(wb, "Микроклимат"));


        // Ориентация страницы: альбомная
        PrintSetup ps = sh.getPrintSetup();
        ps.setLandscape(true);
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        sh.setRepeatingRows(new CellRangeAddress(4, 4, 0, 21)); // 5-я строка, столбцы A..V
        // ===== ширина столбцов (A..V) в пикселях =====
        int[] px = {33,200,28,31,30,34,22,31,40,99,34,14,31,40,34,20,31,40,23,23,23,86};
        for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

        // ===== высота строк (в поинтах) =====
        ensureRow(sh, 0).setHeightInPoints(15);  // ~20 px
        ensureRow(sh, 1).setHeightInPoints(15);  // ~20 px
        ensureRow(sh, 2).setHeightInPoints(48);  // 3-я
        ensureRow(sh, 3).setHeightInPoints(121); // 4-я
        ensureRow(sh, 4).setHeightInPoints(15);  // ~20 px (5-я)

        PageBreakHelper pageBreakHelper = new PageBreakHelper(sh, 5);

        // ===== заголовки =====
        put(sh, 0, 0, "15. Результаты измерений параметров микроклимата", S.title);
        put(sh, 1, 0, "15.2.  Показатели микроклимата в помещениях:", S.title);

        // ===== шапка =====
        mergeWithBorder(sh, "A3:A4"); put(sh, 2, 0, "№ п/п", S.head8);
        mergeWithBorder(sh, "B3:B4"); put(sh, 2, 1,
                "Рабочее место, место проведения измерений, цех, участок, наименование профессии или должности",
                S.head8Wrap);

        mergeWithBorder(sh, "C3:C4"); put(sh, 2, 2, "Дата, проведения  измерений", S.head8Vertical);
        mergeWithBorder(sh, "D3:D4"); put(sh, 2, 3, "Категория работ по интенсивности", S.head8Vertical);
        mergeWithBorder(sh, "E3:E4"); put(sh, 2, 4, "Высота от пола, м", S.head8Vertical);

        mergeWithBorder(sh, "F3:I3"); put(sh, 2, 5, "Температура воздуха, ºС", S.head8Wrap);
        mergeWithBorder(sh, "F4:H4"); put(sh, 3, 5, "Измеренная\n(± расширенная неопределенность)", S.head8Wrap);
        put(sh, 3, 8,  "Допустимый уровень", S.head8Vertical);

        put(sh, 2, 9,  "Температура поверхностей, ºС", S.head8Wrap);
        put(sh, 3, 9,  "Пол. Измеренная\n(± расширенная неопределенность)", S.head8Wrap);

        mergeWithBorder(sh, "K3:N3"); put(sh, 2, 10, "Результирующая температура,  ºС", S.head8Wrap);
        mergeWithBorder(sh, "K4:M4"); put(sh, 3, 10, "Измеренная\n(± расширенная неопределенность)", S.head8Wrap);
        put(sh, 3, 13, "Допустимый уровень", S.head8Vertical);

        mergeWithBorder(sh, "O3:R3"); put(sh, 2, 14, "Относительная влажность воздуха, %", S.head8Wrap);
        mergeWithBorder(sh, "O4:Q4"); put(sh, 3, 14, "Измеренная\n(± расширенная неопределенность)", S.head8Wrap);
        put(sh, 3, 17, "Допустимый уровень", S.head8Vertical);

        mergeWithBorder(sh, "S3:V3"); put(sh, 2, 18, "Скорость движения воздуха,  м/с", S.head8Wrap);
        mergeWithBorder(sh, "S4:U4"); put(sh, 3, 18, "Измеренная\n(± расширенная неопределенность)", S.head8Wrap);
        put(sh, 3, 21, "Допустимый уровень", S.head8Vertical);

        // ===== строка 5 (номера граф) =====
        mergeWithBorder(sh, "F5:H5");
        mergeWithBorder(sh, "K5:M5");
        mergeWithBorder(sh, "O5:Q5");
        mergeWithBorder(sh, "S5:U5");
        put(sh, 4, 0,  1, S.centerBorder);
        put(sh, 4, 1,  2, S.centerBorder);
        put(sh, 4, 2,  3, S.centerBorder);
        put(sh, 4, 3,  4, S.centerBorder);
        put(sh, 4, 4,  5, S.centerBorder);
        put(sh, 4, 5,  6, S.centerBorder); // F5-H5
        put(sh, 4, 8,  7, S.centerBorder);
        put(sh, 4, 9,  8, S.centerBorder);
        put(sh, 4, 10, 9, S.centerBorder); // K5-M5
        put(sh, 4, 13,10, S.centerBorder);
        put(sh, 4, 14,11, S.centerBorder); // O5-Q5
        put(sh, 4, 17,12, S.centerBorder);
        put(sh, 4, 18,13, S.centerBorder); // S5-U5
        put(sh, 4, 21,14, S.centerBorder);

        addRegionBorders(sh, 2, 4, 0, 21); // A3:V5

        // ===== генераторы =====
        Random rng = new Random();
        Map<Space, Double> spaceOBase = new IdentityHashMap<>();
        Map<Room,  Double> roomOBase  = new IdentityHashMap<>();

        // ===== данные =====
        int row = 5; // начинаем вывод с 6-й строки (0-based)
        int seq = 1;

        DateTimeFormatter dateFmt = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        String today = LocalDate.now().format(dateFmt);

        // секция -> этаж -> список выбранных комнат (именно MicroclimateSelected)
        Map<Integer, Map<Floor, List<Room>>> bySection = new LinkedHashMap<>();
        for (Floor floor : building.getFloors()) {
            if (floor == null) continue;
            if (sectionIndex >= 0 && floor.getSectionIndex() != sectionIndex) continue;

            for (Space space : floor.getSpaces()) {
                if (space == null) continue;

                List<Room> selected = new ArrayList<>();
                for (Room r : space.getRooms()) {
                    if (r != null && r.isMicroclimateSelected()) selected.add(r);
                }
                if (selected.isEmpty()) continue;

                bySection.computeIfAbsent(floor.getSectionIndex(), k -> new LinkedHashMap<>())
                        .computeIfAbsent(floor, k -> new ArrayList<>())
                        .addAll(selected);
            }
        }

        List<Integer> secOrder = new ArrayList<>(bySection.keySet());
        secOrder.sort(Comparator.naturalOrder());

        for (Integer secIdx : secOrder) {
            Map<Floor, List<Room>> byFloor = bySection.get(secIdx);

            // этажи по возрастанию номера
            List<Floor> floors = new ArrayList<>(byFloor.keySet());
            floors.sort(Comparator.comparingInt(MicroclimateExcelExporter::parseFloorNumSafe));

            for (Floor f : floors) {
                // заголовок (A:V объединено); если секция одна — не печатаем её имя
                String title = floorTitleForExcel(f); // ровно то, что пользователь ввёл (номер этажа)
                pageBreakHelper.ensureSpace(row, 1);
                mergeWithBorder(sh, "A" + (row + 1) + ":V" + (row + 1));
                put(sh, row, 0, title, S.sectionHeader);
                row++;
                pageBreakHelper.consume(1);

                // сгруппировать по помещению
                Map<Space, List<Room>> bySpace = new LinkedHashMap<>();
                for (Room r : byFloor.get(f)) {
                    Space sp = findRoomSpace(f, r);
                    if (sp == null) continue;
                    bySpace.computeIfAbsent(sp, k -> new ArrayList<>()).add(r);
                }

                for (Map.Entry<Space, List<Room>> g : bySpace.entrySet()) {
                    Space space = g.getKey();
                    List<Room> rooms = g.getValue();

                    boolean isResidential = isApartment(space);
                    boolean isOffice      = isOfficeOrPublic(space);

                    // базовый O для квартиры (Space)
                    spaceOBase.computeIfAbsent(space, s -> rnd(rng, 34.0, 42.0));

                    for (Room room : rooms) {
                        // базовый O для комнаты (внутри ±0.5 от квартиры)
                        roomOBase.computeIfAbsent(room, r -> clamp(rnd(rng, spaceOBase.get(space) - 0.5, spaceOBase.get(space) + 0.5), 34.0, 42.0));

                        // 1) центральная точка (3 строки по высотам)
                        pageBreakHelper.ensureSpace(row, 3);
                        row = emitBlock(sh, row, seq++, building, f, space, room,
                                Position.CENTER, today, isResidential, isOffice, mode, S, rng, spaceOBase, roomOBase);
                        pageBreakHelper.consume(3);

                        // 2) у наружной стены — по числу стен (0..N)
                        int walls = externalWallsCount(room);
                        if (isOffice) {
                            if (walls >= 1) {
                                pageBreakHelper.ensureSpace(row, 3);
                                row = emitBlock(sh, row, seq++, building, f, space, room,
                                        Position.NEAR_WALL, today, isResidential, isOffice, mode, S, rng, spaceOBase, roomOBase);
                                pageBreakHelper.consume(3);
                            }
                        } else {
                            for (int i = 0; i < Math.min(walls, 2); i++) {
                                pageBreakHelper.ensureSpace(row, 3);
                                row = emitBlock(sh, row, seq++, building, f, space, room,
                                        Position.NEAR_WALL, today, isResidential, isOffice, mode, S, rng, spaceOBase, roomOBase);
                                pageBreakHelper.consume(3);
                            }
                        }
                    }
                }
            }
        }

        // после 5-й — все строки по ~20 px
        for (int r = 5; r <= Math.max(5, sh.getLastRowNum()); r++) {
            ensureRow(sh, r).setHeightInPoints(15f);
        }
    }

    /* ======== ВСЁ НИЖЕ — БЕЗ ИЗМЕНЕНИЙ (утилиты, стили, генерация) ======== */

    private enum Position { CENTER, NEAR_WALL }

    /** вывод одного блока (3 строки) + генерация значений */
    private static int emitBlock(
            Sheet sh, int row, int seq, Building building, Floor f, Space space, Room room,
            Position pos, String date, boolean isResidential, boolean isOffice, TemperatureMode mode,
            Styles S, Random rng, Map<Space, Double> spaceOBase, Map<Room, Double> roomOBase) {

        int start = row;

        // A — № п/п (объединяем 3 строки)
        mergeWithBorder(sh, "A" + (start + 1) + ":A" + (start + 3));
        put(sh, start, 0, seq, S.centerBorder);

        // B — место измерения (3 строки) — БЕЗ запятой перед скобками
        String roomName  = room.getName() != null ? room.getName() : "";
        String spaceName = spaceDisplayName(space);
        String postfix   = (pos == Position.CENTER) ? "(в центре помещения)" : "(0,5 метра от наружной стены)";

        String base = isOffice ? nonEmpty(roomName) : joinComma(spaceName, roomName);
        String where = base + (postfix.isBlank() ? "" : " " + postfix);

        boolean highlightBlue = externalWallsCount(room) > 1;
        mergeWithBorder(sh, "B" + (start + 1) + ":B" + (start + 3));
        put(sh, start, 1, where, highlightBlue ? S.textLeftBorderWrapBlue : S.textLeftBorderWrap);

        // C — ДАТА: объединено, поворот 90° (снизу-вверх)
        mergeWithBorder(sh, "C" + (start + 1) + ":C" + (start + 3));
        put(sh, start, 2, date, S.vertCenterBorderUp);

        // D — категория: офис — «3а», иначе «-»
        String cat = isOffice ? "3а" : "-";
        mergeWithBorder(sh, "D" + (start + 1) + ":D" + (start + 3));
        put(sh, start, 3, cat, S.centerBorder);

        // E — высоты от пола
        double[] heights = isOffice ? new double[]{1.7, 0.6, 0.1} : new double[]{1.5, 0.6, 0.1};
        for (int i = 0; i < 3; i++) put(sh, start + i, 4, heights[i], S.num1);

        // I (8), N (13): допуски
        String allowI = isOffice ? "19-23" : tSummer(room.getName());
        String allowN = isOffice ? "19-22" : tWinter(room.getName());
        for (int r = 0; r < 3; r++) {
            put(sh, start + r, 8,  allowI, S.centerBorder);
            put(sh, start + r, 13, allowN, S.centerBorder);
        }

        // J — объединить 3 строки, прочерк
        mergeWithBorder(sh, "J" + (start + 1) + ":J" + (start + 3));
        put(sh, start, 9, "-", S.centerBorder);

        // ====== ГЕНЕРАЦИЯ ДАННЫХ ======

        // --- F/G/H ---
        double[] baseFVals = genFTriplet(rng, isOffice);
        double[] fVals = applyAirTempOffset(baseFVals, mode);
        for (int i = 0; i < 3; i++) {
            put(sh, start + i, 5, fVals[i], S.centerNoRightNum1); // F 0.0
            put(sh, start + i, 6, "±",      S.plusMinusTB);       // G
            put(sh, start + i, 7, 0.2,      S.centerNoLeftNum1);  // H 0.0
        }

        // --- K/L/M ---
        double[] kVals = genKFromF(rng, baseFVals); // ≥ 19.7; шаги ≤ 0.2
        for (int i = 0; i < 3; i++) {
            put(sh, start + i, 10, kVals[i], S.centerNoRightNum1); // K 0.0
            put(sh, start + i, 11, "±",       S.plusMinusTB);      // L
            put(sh, start + i, 12, 0.6,       S.centerNoLeftNum1); // M 0.0
        }

        // --- O/P/Q ---
        double roomBase = roomOBase.get(room);
        double[] oVals = {
                round1(clamp(rnd(rng, roomBase - 0.1, roomBase + 0.1), 34.0, 42.0)),
                round1(clamp(rnd(rng, roomBase - 0.1, roomBase + 0.1), 34.0, 42.0)),
                round1(clamp(rnd(rng, roomBase - 0.1, roomBase + 0.1), 34.0, 42.0))
        };
        for (int i = 0; i < 3; i++) {
            put(sh, start + i, 14, oVals[i], S.centerNoRightNum1); // O 0.0
            put(sh, start + i, 15, "±",       S.plusMinusTB);      // P
            put(sh, start + i, 16, 3.5,       S.centerNoLeftNum1); // Q 0.0
        }

        // R — влажность допустимая
        String allowR;
        if (isOffice) {
            allowR = "30-60";
        } else if (isKitchen(room.getName()) || isBathroomToilet(room.getName()) || isBathRoomProper(room.getName())) {
            allowR = "-";
        } else if (isLiving(room.getName())) {
            allowR = "30-60";
        } else {
            allowR = "30-60";
        }
        for (int r = 0; r < 3; r++) put(sh, start + r, 17, allowR, S.centerBorder);

        // S:U — объединяем, пишем "<0,1"
        for (int r = 0; r < 3; r++) {
            String addr = "S" + (start + r + 1) + ":U" + (start + r + 1);
            mergeWithBorder(sh, addr);
            put(sh, start + r, 18, "<0,1", S.centerBorder);
        }

        // V — доп. уровень
        String allowV = isOffice ? "не более 0,3" : "не более 0,2";
        for (int r = 0; r < 3; r++) put(sh, start + r, 21, allowV, S.centerBorder);

        // добить стилями пустые (на всякий случай)
        for (int r = 0; r < 3; r++) {
            for (int c = 5; c <= 21; c++) {
                Cell cell = cell(ensureRow(sh, start + r), c);
                if (cell.getCellStyle() == null || cell.getCellStyle().getIndex() == 0) {
                    cell.setCellStyle(S.centerBorder);
                }
            }
            ensureRow(sh, start + r).setHeightInPoints(15);
        }

        return start + 3;
    }
    private static String floorTitleForExcel(Floor f) {
        if (f == null) return "";
        String num = f.getNumber();
        if (num != null && !num.isBlank()) return num.trim();   // «1 этаж», «подвал», «техэтаж» и т. п.
        String nm = f.getName();
        return (nm != null) ? nm.trim() : "";
    }

    // ===== генераторы значений =====
    private static double[] genFTriplet(Random rng, boolean office) {
        double lo3 = office ? 20.0 : 20.4;
        double hi3 = office ? 20.5 : 20.7;
        double topMin = office ? 20.0 : 20.4;
        double topMax = office ? 21.1 : 21.4;

        double r3 = rnd(rng, lo3, hi3);
        double r2 = r3 + rnd(rng, 0.0, 0.3);
        double r1 = r2 + rnd(rng, 0.0, 0.3);

        r1 = clamp(r1, topMin, topMax);
        r2 = clamp(r2, topMin, topMax);
        r3 = clamp(r3, topMin, topMax);

        return new double[]{ round1(r1), round1(r2), round1(r3) }; // строки 1,2,3
    }

    private static double[] genKFromF(Random rng, double[] fVals) {
        double[] k = new double[3];
        k[2] = round1(Math.max(19.7, fVals[2] - rnd(rng, 0.65, 1.0)));
        double k2 = round1(Math.max(19.7, fVals[1] - rnd(rng, 0.65, 1.0)));
        if (Math.abs(k2 - k[2]) > 0.2) k2 = round1(k[2] + Math.signum(k2 - k[2]) * 0.2);
        k[1] = k2;
        double k1 = round1(Math.max(19.7, fVals[0] - rnd(rng, 0.65, 1.0)));
        if (Math.abs(k1 - k[1]) > 0.2) k1 = round1(k[1] + Math.signum(k1 - k[1]) * 0.2);
        k[0] = k1;
        for (int i = 0; i < 3; i++) k[i] = Math.max(19.7, k[i]);
        return k;
    }
    private static double[] applyAirTempOffset(double[] base, TemperatureMode mode) {
        double offset = mode != null ? mode.getOffset() : 0.0;
        double[] adjusted = new double[base.length];
        for (int i = 0; i < base.length; i++) {
            adjusted[i] = round1(base[i] + offset);
        }
        return adjusted;
    }

    // ===== утилиты модели/текста =====
    private static String sectionOnlyName(Building b, int idx) {
        try {
            if (b.getSections() != null && b.getSections().size() <= 1) return "";
            Section s = b.getSections().get(Math.max(0, idx));
            return s != null && s.getName() != null ? s.getName().trim() : "";
        } catch (Exception e) { return ""; }
    }

    private static String spaceDisplayName(Space s) {
        String id = (s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        return id.isBlank() ? "Помещение" : id;
    }
    private static Space findRoomSpace(Floor f, Room r) {
        for (Space s : f.getSpaces()) if (s.getRooms().contains(r)) return s;
        return null;
    }
    private static int parseFloorNumSafe(Floor f) {
        try {
            String s = (f.getNumber() != null) ? f.getNumber() : f.getName();
            if (s == null) return 0;
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("(-?\\d+)").matcher(s);
            return m.find() ? Integer.parseInt(m.group(1)) : 0;
        } catch (Exception e) { return 0; }
    }
    private static boolean isApartment(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("APART");
    }
    private static boolean isOfficeOrPublic(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("OFFICE") || name.contains("PUBLIC");
    }
    private static int externalWallsCount(Room r) {
        Integer v = (r != null) ? r.getExternalWallsCount() : null;
        return (v == null) ? 0 : Math.max(0, v);
    }
    private static String joinComma(String... parts) {
        List<String> xs = new ArrayList<>();
        for (String p : parts) if (p != null && !p.isBlank()) xs.add(p.trim());
        return String.join(", ", xs);
    }
    private static String joinNotEmpty(String... parts) { return joinComma(parts); }
    private static String nonEmpty(String s) { return (s == null) ? "" : s.trim(); }

    private static boolean isKitchen(String n) {
        if (n == null) return false;
        String s = n.toLowerCase(Locale.ROOT);
        return s.contains("кух");
    }
    private static boolean isBathroomToilet(String n) {
        if (n == null) return false;
        String s = n.toLowerCase(Locale.ROOT);
        return s.contains("сануз") || s.contains("сан уз") || s.contains("с/у") || s.equals("су") || s.contains("туалет");
    }
    private static boolean isCombinedBathroom(String n) {
        if (n == null) return false;
        String s = n.toLowerCase(Locale.ROOT);
        return (s.contains("совмещ") && (s.contains("сануз") || s.contains("с/у") || s.equals("су")));
    }
    private static boolean isBathRoomProper(String n) {
        if (n == null) return false;
        String s = n.toLowerCase(Locale.ROOT);
        return s.contains("ванн") || s.contains("душ");
    }
    private static boolean isLiving(String n) {
        if (n == null) return false;
        String s = n.toLowerCase(Locale.ROOT);
        return (s.contains("жила") || s.contains("комната") || s.contains("спальн") || s.contains("гостиная")) && !isKitchen(n);
    }
    private static String tSummer(String roomName) {
        if (isKitchen(roomName))            return "18-26";
        if (isBathroomToilet(roomName))     return "18-26";
        if (isBathRoomProper(roomName))     return "18-26";
        if (isLiving(roomName))             return "20-24";
        return "-";
    }
    private static String tWinter(String roomName) {
        if (isKitchen(roomName))            return "17-26";
        if (isCombinedBathroom(roomName))   return "17-26";
        if (isBathRoomProper(roomName))     return "17-26";
        if (isBathroomToilet(roomName))     return "17-25";
        if (isLiving(roomName))             return "19-23";
        return "-";
    }

    private static final class PageBreakHelper {
        private static final double POINTS_PER_INCH = 72.0;
        private static final double A4_HEIGHT_POINTS = 842.0;
        private static final double A4_WIDTH_POINTS = 595.0;

        private final Sheet sheet;
        private final int dataStartRow;
        private final int rowsPerPage;
        private int rowsOnPage = 0;

        private PageBreakHelper(Sheet sheet, int dataStartRow) {
            this.sheet = sheet;
            this.dataStartRow = dataStartRow;
            this.rowsPerPage = calculateRowsPerPage();
        }

        void ensureSpace(int rowIndex, int neededRows) {
            if (rowIndex < dataStartRow || rowsPerPage <= 0 || neededRows > rowsPerPage) {
                return;
            }
            if (rowsOnPage + neededRows > rowsPerPage) {
                sheet.setRowBreak(rowIndex - 1);
                rowsOnPage = 0;
            }
        }

        void consume(int rows) {
            if (rowsPerPage <= 0) return;
            rowsOnPage += rows;
        }

        private int calculateRowsPerPage() {
            double rowHeight = 15.0;
            double headerHeight = 0.0;
            for (int i = 0; i < dataStartRow; i++) {
                headerHeight += ensureRow(sheet, i).getHeightInPoints();
            }
            double marginPoints = (sheet.getMargin(Sheet.TopMargin) + sheet.getMargin(Sheet.BottomMargin)) * POINTS_PER_INCH;
            double pageHeight = pageHeightPoints();
            double available = pageHeight - marginPoints - headerHeight;
            if (available <= 0) return 0;
            return (int) Math.floor(available / rowHeight);
        }

        private double pageHeightPoints() {
            PrintSetup ps = sheet.getPrintSetup();
            if (ps != null && ps.getPaperSize() == PrintSetup.A4_PAPERSIZE) {
                return ps.getLandscape() ? A4_WIDTH_POINTS : A4_HEIGHT_POINTS;
            }
            return A4_WIDTH_POINTS;
        }
    }

    // ===== табличные/числовые утилиты =====
    private static Row ensureRow(Sheet sh, int r0) {
        Row r = sh.getRow(r0);
        return (r != null) ? r : sh.createRow(r0);
    }
    private static Cell cell(Row r, int c0) {
        Cell c = r.getCell(c0);
        return (c != null) ? c : r.createCell(c0);
    }
    private static void put(Sheet sh, int r0, int c0, String val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }
    private static void put(Sheet sh, int r0, int c0, int val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }
    private static void put(Sheet sh, int r0, int c0, double val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void setColWidthPx(Sheet sh, int col, int px) {
        sh.setColumnWidth(col, pixel2WidthUnits(px));
    }
    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem   = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static void addRegionBorders(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress region = new CellRangeAddress(r1, r2, c1, c2);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);
    }
    private static void mergeWithBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(r);
        RegionUtil.setBorderTop(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, r, sh);
    }

    private static double rnd(Random rng, double a, double b) { return a + (b - a) * rng.nextDouble(); }
    private static double clamp(double v, double a, double b) { return Math.max(a, Math.min(b, v)); }
    private static double round1(double v) { return Math.round(v * 10.0) / 10.0; }

    private static String uniqueName(Workbook wb, String base) {
        String name = base; int i = 2;
        while (wb.getSheet(name) != null) name = base + " (" + (i++) + ")";
        return name;
    }

    // ===== стили =====
    private static final class Styles {
        final Font arial10, arial8;
        final CellStyle title;
        // шапка (8 пт)
        final CellStyle head8, head8Wrap, head8Vertical;
        // данные
        final CellStyle centerBorder, centerBorderNum1, textLeftBorderWrap, textLeftBorderWrapBlue, num1, vertCenterBorderUp, sectionHeader;
        // «±» и соседи без боковых линий
        final CellStyle plusMinusTB;       // «±» только верх/низ
        final CellStyle centerNoLeft, centerNoRight;
        final CellStyle centerNoLeftNum1, centerNoRightNum1;

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();

            arial10 = wb.createFont(); arial10.setFontName("Arial"); arial10.setFontHeightInPoints((short)10);
            arial8  = wb.createFont(); arial8.setFontName("Arial");  arial8.setFontHeightInPoints((short)8);

            title = wb.createCellStyle(); title.setFont(arial10);

            head8 = wb.createCellStyle();
            head8.setAlignment(HorizontalAlignment.CENTER);
            head8.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(head8); head8.setFont(arial8);

            head8Wrap = wb.createCellStyle(); head8Wrap.cloneStyleFrom(head8); head8Wrap.setWrapText(true);
            head8Vertical = wb.createCellStyle(); head8Vertical.cloneStyleFrom(head8); head8Vertical.setRotation((short)90);

            centerBorder = wb.createCellStyle();
            centerBorder.setAlignment(HorizontalAlignment.CENTER);
            centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(centerBorder); centerBorder.setFont(arial10);

            centerBorderNum1 = wb.createCellStyle();
            centerBorderNum1.cloneStyleFrom(centerBorder);
            centerBorderNum1.setDataFormat(fmt.getFormat("0.0"));

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap); textLeftBorderWrap.setFont(arial10);

            textLeftBorderWrapBlue = wb.createCellStyle();
            textLeftBorderWrapBlue.cloneStyleFrom(textLeftBorderWrap);
            textLeftBorderWrapBlue.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            textLeftBorderWrapBlue.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            num1 = wb.createCellStyle();
            num1.cloneStyleFrom(centerBorder);
            num1.setDataFormat(fmt.getFormat("0.0"));

            // вертикально (90°)
            vertCenterBorderUp = wb.createCellStyle();
            vertCenterBorderUp.cloneStyleFrom(centerBorder);
            vertCenterBorderUp.setRotation((short)90);

            sectionHeader = wb.createCellStyle();
            sectionHeader.setAlignment(HorizontalAlignment.CENTER);
            sectionHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(sectionHeader); sectionHeader.setFont(arial10);

            plusMinusTB = wb.createCellStyle();
            plusMinusTB.setAlignment(HorizontalAlignment.CENTER);
            plusMinusTB.setVerticalAlignment(VerticalAlignment.CENTER);
            plusMinusTB.setBorderTop(BorderStyle.THIN);
            plusMinusTB.setBorderBottom(BorderStyle.THIN);
            plusMinusTB.setBorderLeft(BorderStyle.NONE);
            plusMinusTB.setBorderRight(BorderStyle.NONE);
            plusMinusTB.setFont(arial10);

            centerNoLeft = wb.createCellStyle();
            centerNoLeft.cloneStyleFrom(centerBorder);
            centerNoLeft.setBorderLeft(BorderStyle.NONE);

            centerNoRight = wb.createCellStyle();
            centerNoRight.cloneStyleFrom(centerBorder);
            centerNoRight.setBorderRight(BorderStyle.NONE);

            centerNoLeftNum1 = wb.createCellStyle();
            centerNoLeftNum1.cloneStyleFrom(centerNoLeft);
            centerNoLeftNum1.setDataFormat(fmt.getFormat("0.0"));

            centerNoRightNum1 = wb.createCellStyle();
            centerNoRightNum1.cloneStyleFrom(centerNoRight);
            centerNoRightNum1.setDataFormat(fmt.getFormat("0.0"));
        }

        private static void setAllBorders(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }
}
