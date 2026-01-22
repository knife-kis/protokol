package ru.citlab24.protokol.tabs.modules.lighting;

import org.apache.poi.ss.usermodel.*;
import ru.citlab24.protokol.export.PrintSetupUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.Month;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;

public final class LightingExcelExporter {

    private LightingExcelExporter() {}

    // === ТОНКАЯ ОБЁРТКА: создаёт книгу, вызывает appendToWorkbook(...), предлагает «Сохранить» ===
    public static void export(Building building, int sectionIndex, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = new XSSFWorkbook()) {
            appendToWorkbook(building, sectionIndex, wb);

            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Освещение_естественное.xlsx"));
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
            JOptionPane.showMessageDialog(parent, "Ошибка экспорта: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }

    // === НОВОЕ: дописывает лист(ы) освещения в уже открытую книгу, ничего не сохраняет/не закрывает ===
    public static void appendToWorkbook(Building building, int sectionIndex, Workbook wb) {
        if (building == null || wb == null) return;
        buildLightingSheets(building, sectionIndex, wb);
    }

    // === ВЕСЬ КОД ПОСТРОЕНИЯ ЛИСТА ПЕРЕНЕСЁН СЮДА (без диалога «Сохранить») ===
    private static void buildLightingSheets(Building building, int sectionIndex, Workbook wb) {
        Styles S = new Styles(wb);
        Sheet sh = wb.createSheet("КЕО");
        setupLandscapePage(sh);
        sh.setRepeatingRows(CellRangeAddress.valueOf("5:5"));
        setupPage(sh);

        // ===== ширины столбцов =====
        int[] px = { 49, 250, 87, 47, 47, 47, 44, 18, 44, 44, 18, 44, 44, 18, 44, 42, 40, 20, 40 };
        for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

        // ===== высота строк =====
        ensureRow(sh, 0).setHeightInPoints(16);
        ensureRow(sh, 1).setHeightInPoints(4);
        ensureRow(sh, 2).setHeightInPoints(51);
        ensureRow(sh, 3).setHeightInPoints(160);
        ensureRow(sh, 4).setHeightInPoints(16);

        // ===== A1 =====
        put(sh, 0, 0, "18.2 Естественное освещение:", S.title);

        // ===== Шапка =====
        mergeWithBorder(sh, "A3:A4"); put(sh, 2, 0, "№ п/п", S.headCenterBorder);
        mergeWithBorder(sh, "B3:B4"); put(sh, 2, 1, "Наименование места\nпроведения измерений", S.headCenterBorderWrap);
        mergeWithBorder(sh, "C3:C4"); put(sh, 2, 2, "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м", S.headVertical);

        mergeWithBorder(sh, "D3:F3");
        put(sh, 2, 3, "При верхнем или комбинированном освещении", S.headCenterBorderWrap);
        put(sh, 3, 3, "Освещенность внутри помещения, лк", S.headVertical);
        put(sh, 3, 4, "Наружная освещенность, лк", S.headVertical);
        put(sh, 3, 5, "КЕО, %", S.headVertical);

        mergeWithBorder(sh, "G3:P3"); put(sh, 2, 6, "При боковом освещении", S.headCenterBorderWrap);
        mergeWithBorder(sh, "G4:I4"); put(sh, 3, 6, "Освещенность внутри помещения ± расширенная неопределенность, лк", S.headVertical);
        mergeWithBorder(sh, "J4:L4"); put(sh, 3, 9, "Наружная освещенность ± расширенная неопределенность, лк", S.headVertical);
        mergeWithBorder(sh, "M4:O4"); put(sh, 3, 12, "КЕО ± расширенная неопределенность, %", S.headVertical);
        put(sh, 3, 15, "Допустимое значение КЕО, %", S.headVertical);

        mergeWithBorder(sh, "Q3:S4");
        put(sh, 2, 16, "Неравномерность естественного освещения ± расширенная неопределенность", S.headVertical);

        // Нумерация (ряд 5)
        put(sh, 4, 0, 1, S.headCenterBorder);
        put(sh, 4, 1, 2, S.headCenterBorder);
        put(sh, 4, 2, 3, S.headCenterBorder);
        put(sh, 4, 3, 4, S.headCenterBorder);
        put(sh, 4, 4, 5, S.headCenterBorder);
        put(sh, 4, 5, 6, S.headCenterBorder);
        mergeWithBorder(sh, "G5:I5");  put(sh, 4, 6, 7, S.headCenterBorder);
        mergeWithBorder(sh, "J5:L5");  put(sh, 4, 9, 8, S.headCenterBorder);
        mergeWithBorder(sh, "M5:O5");  put(sh, 4, 12, 9, S.headCenterBorder);
        put(sh, 4, 15, 10, S.headCenterBorder);
        mergeWithBorder(sh, "Q5:S5");  put(sh, 4, 16, 11, S.headCenterBorder);

        addRegionBorders(sh, 2, 4, 0, 18); // рамка А3:S5

        // ===== Данные =====
        int row = 5;   // 0-based вывод
        int seq = 1;

        Month m = LocalDate.now().getMonth();
        IntRange jSeason = seasonalJRange(m);

        // собираем Rooms, сгруппировано по Space (квартире/офису) в порядке обхода
        // ВНИМАНИЕ: как и раньше — берём все секции (sectionIndex не фильтруем), т.к. ранее было -1
        List<Entry> flat = collectEntries(building, -1);
        LinkedHashMap<Space, List<Entry>> bySpace = new LinkedHashMap<>();
        for (Entry e : flat) bySpace.computeIfAbsent(e.space, k -> new ArrayList<>()).add(e);

        Integer prevLastJOfApartment = null; // «якорь» J между КВАРТИРАМИ

        for (Map.Entry<Space, List<Entry>> group : bySpace.entrySet()) {
            Space space = group.getKey();
            List<Entry> rooms = group.getValue();

            boolean isResidential = isResidentialSpace(space);
            boolean isOfficePublic = isOfficeOrPublic(space);

            List<RoomBlock> blocks = new ArrayList<>();
            Integer localPrevAnchorForFirstBlock = (isResidential ? prevLastJOfApartment : null);

            // --- комнаты данного помещения ---
            for (int roomIdx = 0; roomIdx < rooms.size(); roomIdx++) {
                Entry e = rooms.get(roomIdx);
                boolean isKitchen = isKitchenName(e.room.getName());

                int blockStart = row;
                int blockEnd = row + 4;
                for (int r = blockStart; r <= blockEnd; r++) ensureRow(sh, r).setHeightInPoints(16);

                // B — объединяем на 5 строк + текст
                mergeWithBorder(sh, "B" + (blockStart + 1) + ":B" + (blockEnd + 1));
                put(sh, blockStart, 1, buildBLabel(building, e), S.centerBorderWrap);

                // Q..S — единый прямоугольник 3x5 с «-»
                String qRange = "Q" + (blockStart + 1) + ":S" + (blockEnd + 1);
                mergeWithBorder(sh, qRange);
                put(sh, blockStart, 16, "-", S.center);

                // J-ряд (по сезону, |Δ| ≤ 60, округление до десятков)
                int[] Jvals = generateJSeries(5, jSeason.min, jSeason.max,
                        (roomIdx == 0 ? localPrevAnchorForFirstBlock : null), 150);

                double[] Mvals = new double[5];
                double[] Gvals = new double[5];
                Double prevMRow = null;

                // Границы М и мин. разницы по типу помещения
                double[][] M_RANGES_RES = {
                        {4.88, 5.55},
                        {3.55, 4.54},
                        {1.01, 1.45},
                        {0.71, 0.91},
                        {0.59, 0.68}
                };
                double[] M_DIFF_RES = {0.0, 0.62, 0.0, 0.24, 0.12};

                double[][] M_RANGES_OFF = {
                        {4.89, 5.85},
                        {4.25, 5.44},
                        {1.85, 2.49},
                        {1.55, 1.89},
                        {1.17, 1.45}
                };
                double[] M_DIFF_OFF = {0.0, 0.35, 0.0, 0.25, 0.25};

                for (int i = 0; i < 5; i++) {
                    int r = row + i;
                    int R = r + 1;
                    Row rr = ensureRow(sh, r);

                    // A — №
                    Cell A = cell(rr, 0); A.setCellValue(seq++); A.setCellStyle(S.centerBorder);

                    // C — Г-0,8 / Г-0,0
                    String cVal = (isOfficePublic ? "Г-0,8" : "Г-0,0");
                    Cell C = cell(rr, 2); C.setCellValue(cVal); C.setCellStyle(S.centerBorder);

                    // D,E,F — прочерк
                    for (int c = 3; c <= 5; c++) {
                        Cell x = cell(rr, c); x.setCellValue("-"); x.setCellStyle(S.centerBorder);
                    }

                    // --- J/K/L ---
                    int Jv = Jvals[i];
                    Cell J = cell(rr, 9);  J.setCellValue(Jv);                        J.setCellStyle(S.num0NoRight);
                    Cell K = cell(rr, 10); K.setCellValue("±");                       K.setCellStyle(S.pmTopBottom);
                    Cell L = cell(rr, 11); L.setCellFormula("J" + R + "*0.08*2/POWER(3,0.5)"); L.setCellStyle(S.num0NoLeft);

                    // --- G/H/I ---
                    IntRange kRange;
                    int gInt;
                    if (isOfficePublic) {
                        kRange = kRangeForM(Jv, M_RANGES_OFF[i][0], M_RANGES_OFF[i][1]);
                        gInt = chooseKWithDiff(Jv, kRange, prevMRow, M_DIFF_OFF[i]);
                    } else { // жилые
                        kRange = kRangeForM(Jv, M_RANGES_RES[i][0], M_RANGES_RES[i][1]);
                        gInt = chooseKWithDiff(Jv, kRange, prevMRow, M_DIFF_RES[i]);
                    }
                    if (gInt < kRange.min) gInt = kRange.min;
                    if (gInt > kRange.max) gInt = kRange.max;

                    double Gv = gInt;
                    if (Gv < 100 && gInt < kRange.max) {
                        int decimals = (Gv >= 10) ? 1 : 2;
                        Gv = addDecimalNoise(gInt, decimals);
                        if (Gv > kRange.max) Gv = kRange.max;
                    }
                    Gvals[i] = Gv;

                    // формат G и I в зависимости от величины G
                    Cell G = cell(rr, 6);
                    Cell H = cell(rr, 7);
                    Cell I = cell(rr, 8);

                    G.setCellValue(Gv);
                    H.setCellValue("±"); H.setCellStyle(S.pmTopBottom);

                    if (Gv >= 100) {
                        G.setCellStyle(S.num0NoRight);
                        I.setCellFormula("G" + R + "*0.08*2/POWER(3,0.5)"); I.setCellStyle(S.num0NoLeft);
                    } else if (Gv >= 10) {
                        G.setCellStyle(S.num1NoRight);
                        I.setCellFormula("G" + R + "*0.08*2/POWER(3,0.5)"); I.setCellStyle(S.num1NoLeft);
                    } else {
                        G.setCellStyle(S.num2NoRight);
                        I.setCellFormula("G" + R + "*0.08*2/POWER(3,0.5)"); I.setCellStyle(S.num2NoLeft);
                    }

                    // --- M/N/O ---
                    Cell Mv = cell(rr, 12); Mv.setCellFormula("G" + R + "*100/J" + R); Mv.setCellStyle(S.num2NoRight);
                    Cell N  = cell(rr, 13); N.setCellValue("±");                       N.setCellStyle(S.pmTopBottom);
                    Cell O  = cell(rr, 14);
                    O.setCellFormula("2*M" + R + "*SQRT(POWER(L" + R + "/2/J" + R + ",2)+POWER(I" + R + "/2/G" + R + ",2))");
                    O.setCellStyle(S.num2NoLeft);

                    // P — для офисов ставим 1,00 сразу
                    if (isOfficePublic) {
                        Cell P = cell(rr, 15); P.setCellValue(1.00); P.setCellStyle(S.num2);
                    }

                    double mApprox = (Gv * 100.0) / Jv;
                    Mvals[i] = mApprox;
                    prevMRow = mApprox;
                }

                RoomBlock b = new RoomBlock(e, blockStart, Gvals, Jvals, Mvals,
                        !isOfficePublic && isResidential, isOfficePublic, isKitchen);
                blocks.add(b);

                // если это был первый блок квартиры и у нас был «якорь», дальше внутри квартиры не ограничиваем старт
                localPrevAnchorForFirstBlock = null;

                row += 5;
            } // конец перебора комнат помещения

            // ===== P-логика для жилых =====
            if (isResidential && !isOfficePublic) {
                // кухни: 3-я строка = 0.50; 1,2,4,5 — «-»
                for (RoomBlock b : blocks) {
                    if (!b.isResidential || !b.isKitchen) continue;
                    setPDashes(sh, b.startRow, S);
                    setPvalue(sh, b.startRow + 2, 0.50, S); // 3-я строка
                }
                // остальные комнаты (не кухни)
                List<RoomBlock> rest = new ArrayList<>();
                for (RoomBlock b : blocks) if (b.isResidential && !b.isKitchen) rest.add(b);

                if (rest.size() == 1) {
                    RoomBlock only = rest.get(0);
                    setPDashes(sh, only.startRow, S);
                    setPvalue(sh, only.startRow + 4, 0.50, S);
                } else if (rest.size() >= 2) {
                    RoomBlock target = rest.get(0);
                    double best = min(target.Mvals);
                    for (int i = 1; i < rest.size(); i++) {
                        double mmin = min(rest.get(i).Mvals);
                        if (mmin < best) { best = mmin; target = rest.get(i); }
                    }
                    for (RoomBlock b : rest) {
                        setPDashes(sh, b.startRow, S);
                        if (b == target) setPvalue(sh, b.startRow + 4, 0.50, S);
                        else             setPvalue(sh, b.startRow + 2, 0.50, S);
                    }
                }
            }

            // ===== ОБНОВЛЯЕМ «якорь» J между квартирами =====
            if (isResidential) {
                // берём J последнего блока/5-й строки
                if (!blocks.isEmpty()) {
                    RoomBlock last = blocks.get(blocks.size() - 1);
                    prevLastJOfApartment = last.Jvals[4];
                }
            }
        }

        int noteRow = row;
        Row note = ensureRow(sh, noteRow);
        note.setHeightInPoints(pixelsToPoints(67));
        String noteRange = "A" + (noteRow + 1) + ":S" + (noteRow + 1);
        mergeWithoutBorder(sh, noteRange);
        addRegionTopBorder(sh, noteRange);
        String noteText = buildCustomerDataNote(wb);
        put(sh, noteRow, 0, noteText, S.textLeftItalicTopBorderWrap);

        int footerRow = noteRow + 1;
        Row spacer1 = ensureRow(sh, footerRow++);
        spacer1.setHeightInPoints(pixelsToPoints(8));

        Row measurerRow = ensureRow(sh, footerRow++);
        measurerRow.setHeightInPoints(pixelsToPoints(17));
        mergeWithoutBorder(sh, "A" + (measurerRow.getRowNum() + 1) + ":S" + (measurerRow.getRowNum() + 1));
        put(sh, measurerRow.getRowNum(), 0,
                "Измерения проводил: заведующий лабораторией ____________ Тарновский М.О.",
                S.textLeft);

        Row measurerSignRow = ensureRow(sh, footerRow++);
        mergeWithoutBorder(sh, "A" + (measurerSignRow.getRowNum() + 1) + ":G" + (measurerSignRow.getRowNum() + 1));
        put(sh, measurerSignRow.getRowNum(), 0, "(должность, подпись, Ф.И.О.)", S.center8);

        Row preparedRow = ensureRow(sh, footerRow++);
        mergeWithoutBorder(sh, "A" + (preparedRow.getRowNum() + 1) + ":S" + (preparedRow.getRowNum() + 1));
        put(sh, preparedRow.getRowNum(), 0,
                "Протокол подготовил: заведующий лабораторией ____________ Тарновский М.О.",
                S.textLeft);

        Row preparedSignRow = ensureRow(sh, footerRow++);
        preparedSignRow.setHeightInPoints(pixelsToPoints(20));
        mergeWithoutBorder(sh, "A" + (preparedSignRow.getRowNum() + 1) + ":G" + (preparedSignRow.getRowNum() + 1));
        put(sh, preparedSignRow.getRowNum(), 0, "(должность, подпись, Ф.И.О.)", S.center8);

        Row responsibilityRow = ensureRow(sh, footerRow++);
        responsibilityRow.setHeightInPoints(pixelsToPoints(109));
        mergeWithoutBorder(sh, "A" + (responsibilityRow.getRowNum() + 1) + ":S" + (responsibilityRow.getRowNum() + 1));
        put(sh, responsibilityRow.getRowNum(), 0,
                "Испытательная лаборатория несет ответственность за всю информацию, представленную в протоколе испытаний, за исключением случаев, когда информация предоставляется заказчиком.\n"
                        + "Протокол не должен быть воспроизведен не в полном объеме без разрешения испытательной лаборатории ООО «ЦИТ».\n"
                        + "Результаты относятся только к объектам, прошедшим испытания и предоставленным заказчиком.\n"
                        + "Распределение экземпляров протокола испытаний: два протокола – Заказчику, один протокол – испытательной лаборатории ООО «ЦИТ».",
                S.centerWrap);

        Row spacer2 = ensureRow(sh, footerRow++);
        spacer2.setHeightInPoints(pixelsToPoints(7));
        String bottomBorderRange = "A" + (spacer2.getRowNum() + 1) + ":S" + (spacer2.getRowNum() + 1);
        mergeWithoutBorder(sh, bottomBorderRange);
        addRegionBottomBorder(sh, bottomBorderRange);

        Row endRow = ensureRow(sh, footerRow);
        mergeWithoutBorder(sh, "A" + (endRow.getRowNum() + 1) + ":S" + (endRow.getRowNum() + 1));
        put(sh, endRow.getRowNum(), 0, "Конец протокола испытаний", S.centerBold);
    }

    // ===== ЛОГИКА ДАННЫХ =====

    private static IntRange seasonalJRange(Month m) {
        switch (m) {
            case FEBRUARY, MARCH, APRIL, MAY:      return new IntRange(5800, 6800);
            case JUNE, JULY, AUGUST:               return new IntRange(6700, 7500);
            case SEPTEMBER, OCTOBER, NOVEMBER:     return new IntRange(6000, 6900);
            case DECEMBER, JANUARY:
            default:                                return new IntRange(5400, 6300);
        }
    }

    /** Сгенерировать count значений J в [min..max], для первого элемента можно задать «якорь» и maxDiff, шаги ≤ 60, округление до десятков. */
    private static int[] generateJSeries(int count, int min, int max, Integer startAnchor, int maxStartDiff) {
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        int[] arr = new int[count];

        int start;
        if (startAnchor == null) {
            start = round10(rnd.nextInt(min, max + 1));
        } else {
            int low = Math.max(min, startAnchor - maxStartDiff);
            int high = Math.min(max, startAnchor + maxStartDiff);
            if (low > high) { low = min; high = max; }
            start = round10(rnd.nextInt(low, high + 1));
        }
        arr[0] = start;

        int prev = start;
        for (int i = 1; i < count; i++) {
            int cand; int tries = 0;
            do {
                int delta = rnd.nextInt(-60, 61);
                if (delta == 0) delta = (rnd.nextBoolean() ? 10 : -10);
                cand = clamp(prev + delta, min, max);
                cand = round10(cand);
                tries++;
                if (tries > 200) { cand = prev; break; }
            } while (Math.abs(cand - prev) > 60);
            arr[i] = cand;
            prev = cand;
        }
        return arr;
    }

    /** Диапазон k (k=G) для M в [mMin..mMax] при данном J. */
    private static IntRange kRangeForM(int J, double mMin, double mMax) {
        int kMin = (int) Math.ceil(mMin * J / 100.0);
        int kMax = (int) Math.floor(mMax * J / 100.0);
        if (kMin > kMax) { int t = kMin; kMin = kMax; kMax = t; }
        return new IntRange(kMin, kMax);
    }

    /** Выбор k (G), чтобы |M(k)-prevM| ≥ minDiff, иначе — граница, дающая бОльшую разницу. */
    private static int chooseKWithDiff(int J, IntRange kRange, Double prevM, double minDiff) {
        if (kRange.isEmpty()) return Math.max(0, kRange.min);
        ThreadLocalRandom rnd = ThreadLocalRandom.current();

        if (prevM == null || minDiff <= 0.0) {
            int mid = (kRange.min + kRange.max) / 2;
            int span = Math.max(1, (kRange.max - kRange.min) / 4);
            int raw = mid + rnd.nextInt(-span, span + 1);
            return clamp(raw, kRange.min, kRange.max);
        }
        List<Integer> candidates = new ArrayList<>();
        for (int k = kRange.min; k <= kRange.max; k++) {
            double M = 100.0 * k / J;
            if (Math.abs(M - prevM) >= minDiff) candidates.add(k);
        }
        if (!candidates.isEmpty()) return candidates.get(rnd.nextInt(candidates.size()));
        int k1 = kRange.min, k2 = kRange.max;
        double m1 = 100.0 * k1 / J, m2 = 100.0 * k2 / J;
        return (Math.abs(m1 - prevM) >= Math.abs(m2 - prevM)) ? k1 : k2;
    }

    private static boolean isKitchenName(String name) {
        if (name == null) return false;
        String n = name.toLowerCase(Locale.ROOT);
        return n.contains("кух"); // кухня, кухня-ниша, кухон...
    }

    private static int round10(int v) { return Math.round(v / 10f) * 10; }
    private static int round10Int(int v) { return round10(v); }
    private static int clamp(int v, int min, int max) { return Math.max(min, Math.min(max, v)); }
    private static double min(double[] a) { double m = Double.POSITIVE_INFINITY; for (double x: a) m = Math.min(m, x); return m; }
    private static double addDecimalNoise(int baseValue, int decimals) {
        if (decimals <= 0) return baseValue;
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        int scale = (int) Math.pow(10, decimals);
        int fraction = rnd.nextInt(scale);
        return baseValue + fraction / (double) scale;
    }

    // ===== Табличные утилиты и сбор данных =====

    private static String buildCustomerDataNote(Workbook wb) {
        if (hasVentilationVolumes(wb)) {
            return "Данные, предоставленные заказчиком, за которые он несет ответственность: допустимые уровни в пунктах: 15, 17; нормируемая освещенность в пунктах 18.1, 18.3; допустимые значения в пункте 18.2; объемы помещений в пункте 15.4. Испытательная лаборатория не несет ответственность за данные, предоставленные заказчиком (заявление об ограничении ответственности испытательной лаборатории).";
        }
        return "Данные, предоставленные заказчиком, за которые он несет ответственность: допустимые уровни в пунктах: 15, 17; нормируемая освещенность в пунктах 18.1, 18.3; допустимые значения в пункте 18.2. Испытательная лаборатория не несет ответственность за данные, предоставленные заказчиком (заявление об ограничении ответственности испытательной лаборатории).";
    }

    private static boolean hasVentilationVolumes(Workbook wb) {
        if (wb == null) return false;
        Sheet sheet = findVentilationSheet(wb);
        if (sheet == null) return false;
        DataFormatter formatter = new DataFormatter();
        int lastRow = sheet.getLastRowNum();
        for (int r = 4; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell cell = row.getCell(8);
            if (cell == null) continue;
            String text = formatter.formatCellValue(cell);
            if (text != null && !text.isBlank()) return true;
        }
        return false;
    }

    private static Sheet findVentilationSheet(Workbook wb) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null || sheet.getSheetName() == null) continue;
            String name = sheet.getSheetName().trim().toLowerCase(Locale.ROOT);
            if (name.equals("вентиляция") || name.startsWith("вентиляция ")) return sheet;
        }
        return null;
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

    private static void mergeWithoutBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(r);
    }

    private static void addRegionTopBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        RegionUtil.setBorderTop(BorderStyle.THIN, r, sh);
    }

    private static void addRegionBottomBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        RegionUtil.setBorderBottom(BorderStyle.THIN, r, sh);
    }

    private static void setPDashes(Sheet sh, int startRow, Styles S) {
        for (int i = 0; i < 5; i++) {
            Cell P = cell(ensureRow(sh, startRow + i), 15);
            P.setCellValue("-"); P.setCellStyle(S.centerBorder);
        }
    }

    private static void setPvalue(Sheet sh, int rowIndex, double val, Styles S) {
        Cell P = cell(ensureRow(sh, rowIndex), 15);
        P.setCellValue(val); P.setCellStyle(S.num2);
    }

    private static List<Entry> collectEntries(Building building, int sectionIndex) {
        List<Entry> res = new ArrayList<>();
        for (Floor floor : building.getFloors()) {
            if (floor == null) continue;
            if (sectionIndex >= 0 && floor.getSectionIndex() != sectionIndex) continue;

            for (Space space : floor.getSpaces()) {
                if (space == null) continue;
                for (Room room : space.getRooms()) {
                    if (room != null && room.isSelected()) {
                        res.add(new Entry(floor, space, room));
                    }
                }
            }
        }
        return res;
    }

    private static String sectionNameIfMultiple(Building building, int sectionIndex) {
        if (building == null || building.getSections() == null) return "";
        List<Section> sections = building.getSections();
        if (sections.size() <= 1) return ""; // одна секция — не пишем её в столбец B
        Section s = sections.get(Math.max(0, sectionIndex));
        if (s == null || s.getName() == null || s.getName().isBlank()) return "";
        return s.getName();
    }

    private static String buildBLabel(Building building, Entry e) {
        // Этаж — берём то, что ввели руками (без префиксов типа этажа)
        String floorNameRaw = (e.floor.getName() != null && !e.floor.getName().isBlank())
                ? e.floor.getName()
                : (e.floor.getNumber() != null ? e.floor.getNumber() : "Этаж");
        String floorName = ensureEtaz(cleanedFloorName(floorNameRaw));

        String roomName = (e.room.getName() != null) ? e.room.getName() : "Комната";
        boolean isOffice = isOffice(e.space);

        // если секция одна — возвращаем "", чтобы в столбце B начать сразу с этажа
        String sect = sectionNameIfMultiple(building, e.floor.getSectionIndex());

        if (isOffice) {
            // Офисы: «[Секция,] Этаж, Комната, нормируемая поверхность»
            return joinComma(sect, floorName, roomName) + ", нормируемая поверхность";
        } else {
            // Квартиры/прочие: «[Секция,] Этаж, Помещение, Комната, нормируемая поверхность»
            String spaceName = spaceDisplayName(e.space);
            return joinComma(sect, floorName, spaceName, roomName) + ", нормируемая поверхность";
        }
    }

    private static String joinComma(String... parts) {
        StringBuilder sb = new StringBuilder();
        for (String p : parts) {
            if (p != null && !p.isBlank()) {
                if (sb.length() > 0) sb.append(", ");
                sb.append(p.trim());
            }
        }
        return sb.toString();
    }

    private static boolean isOffice(Space space) {
        if (space == null || space.getType() == null) return false;
        return space.getType() == Space.SpaceType.OFFICE;
    }

    private static String cleanedFloorName(String raw) {
        if (raw == null) return "Этаж";
        String s = raw.replaceFirst("(?iu)^(\\s*)(смешанн(?:ый|ая|ое|ые|ых|ом)?|офисн(?:ый|ая|ое|ые|ых|ом)?|жил(?:ой|ая|ое|ые|ых|ом)?)[\\s,]+", "");
        s = s.replaceAll(" +", " ").trim();
        return s.isEmpty() ? raw : s;
    }
    private static String ensureEtaz(String s) {
        if (s == null) return "Этаж";
        String t = s.trim();
        if (t.isEmpty()) return "Этаж";
        // если пользователь уже написал слово «этаж» — ничего не меняем
        if (t.toLowerCase(java.util.Locale.ROOT).contains("этаж")) return t;
        // если это просто число (возможно отрицательное) — добавим «этаж»
        if (t.matches("-?\\d+")) return t + " этаж";
        return t;
    }

    private static String spaceDisplayName(Space s) {
        if (s == null) return "Помещение";
        String id = s.getIdentifier();
        if (id != null && !id.isBlank()) return id;
        try {
            java.lang.reflect.Method m = s.getClass().getMethod("getName");
            Object v = m.invoke(s);
            if (v != null) {
                String n = String.valueOf(v);
                if (!n.isBlank()) return n;
            }
        } catch (Exception ignored) {}
        return "Помещение";
    }

    private static boolean isResidentialSpace(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("APART"); // APARTMENT
    }

    private static boolean isOfficeOrPublic(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("OFFICE") || name.contains("PUBLIC");
    }

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

    private static void setColWidthPx(Sheet sh, int col, int px) {
        int width = (int) Math.round((px - 5) / 7.0 * 256.0);
        if (width < 0) width = 0;
        sh.setColumnWidth(col, width);
    }

    private static float pixelsToPoints(float px) {
        return px * 72f / 96f;
    }

    // ===== СТИЛИ =====
    private static final class Styles {
        // Шрифты
        final Font arial10;
        final Font arial9;
        final Font arial8;
        final Font arial10Italic;
        final Font arial10Bold;

        // Шапка (Arial 9)
        final CellStyle headCenterBorder;
        final CellStyle headCenterBorderWrap;
        final CellStyle headVertical;

        // Данные (Arial 10)
        final CellStyle title;
        final CellStyle center;
        final CellStyle center8;
        final CellStyle centerBold;
        final CellStyle centerBorder;
        final CellStyle centerBorderWrap;
        final CellStyle centerWrap;
        final CellStyle textLeft;
        final CellStyle textLeftBorderWrap;
        final CellStyle textLeftItalicTopBorderWrap;
        final CellStyle box;
        final CellStyle num0;
        final CellStyle num1;
        final CellStyle num2;
        final CellStyle pmTopBottom;

        // Соседи «±» — убираем примыкающие грани
        final CellStyle num0NoRight, num0NoLeft;
        final CellStyle num1NoRight, num1NoLeft;
        final CellStyle num2NoRight, num2NoLeft;

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();

            arial10 = wb.createFont(); arial10.setFontName("Arial"); arial10.setFontHeightInPoints((short)10);
            arial9  = wb.createFont();  arial9.setFontName("Arial");  arial9.setFontHeightInPoints((short)9);
            arial8  = wb.createFont();  arial8.setFontName("Arial");  arial8.setFontHeightInPoints((short)8);
            arial10Italic = wb.createFont(); arial10Italic.setFontName("Arial"); arial10Italic.setFontHeightInPoints((short)10); arial10Italic.setItalic(true);
            arial10Bold = wb.createFont(); arial10Bold.setFontName("Arial"); arial10Bold.setFontHeightInPoints((short)10); arial10Bold.setBold(true);

            title = wb.createCellStyle();
            title.setAlignment(HorizontalAlignment.LEFT);
            title.setVerticalAlignment(VerticalAlignment.CENTER);
            title.setFont(arial10);

            headCenterBorder = wb.createCellStyle();
            headCenterBorder.setAlignment(HorizontalAlignment.CENTER);
            headCenterBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(headCenterBorder);
            headCenterBorder.setFont(arial9);

            headCenterBorderWrap = wb.createCellStyle();
            headCenterBorderWrap.cloneStyleFrom(headCenterBorder);
            headCenterBorderWrap.setWrapText(true);
            headCenterBorderWrap.setFont(arial9);

            headVertical = wb.createCellStyle();
            headVertical.setAlignment(HorizontalAlignment.CENTER);
            headVertical.setVerticalAlignment(VerticalAlignment.CENTER);
            headVertical.setWrapText(true);
            headVertical.setRotation((short) 90);
            setAllBorders(headVertical);
            headVertical.setFont(arial9);

            center = wb.createCellStyle();
            center.setAlignment(HorizontalAlignment.CENTER);
            center.setVerticalAlignment(VerticalAlignment.CENTER);
            center.setFont(arial10);

            center8 = wb.createCellStyle();
            center8.setAlignment(HorizontalAlignment.CENTER);
            center8.setVerticalAlignment(VerticalAlignment.CENTER);
            center8.setFont(arial8);

            centerBold = wb.createCellStyle();
            centerBold.setAlignment(HorizontalAlignment.CENTER);
            centerBold.setVerticalAlignment(VerticalAlignment.CENTER);
            centerBold.setFont(arial10Bold);

            centerBorder = wb.createCellStyle();
            centerBorder.cloneStyleFrom(center);
            setAllBorders(centerBorder);
            centerBorder.setFont(arial10);

            centerBorderWrap = wb.createCellStyle();
            centerBorderWrap.setAlignment(HorizontalAlignment.CENTER);
            centerBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            centerBorderWrap.setWrapText(true);
            setAllBorders(centerBorderWrap);
            centerBorderWrap.setFont(arial10);

            centerWrap = wb.createCellStyle();
            centerWrap.setAlignment(HorizontalAlignment.CENTER);
            centerWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            centerWrap.setWrapText(true);
            centerWrap.setFont(arial10);

            textLeft = wb.createCellStyle();
            textLeft.setAlignment(HorizontalAlignment.LEFT);
            textLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeft.setFont(arial10);

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap);
            textLeftBorderWrap.setFont(arial10);

            textLeftItalicTopBorderWrap = wb.createCellStyle();
            textLeftItalicTopBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftItalicTopBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftItalicTopBorderWrap.setWrapText(true);
            textLeftItalicTopBorderWrap.setBorderTop(BorderStyle.THIN);
            textLeftItalicTopBorderWrap.setBorderBottom(BorderStyle.NONE);
            textLeftItalicTopBorderWrap.setBorderLeft(BorderStyle.NONE);
            textLeftItalicTopBorderWrap.setBorderRight(BorderStyle.NONE);
            textLeftItalicTopBorderWrap.setFont(arial10Italic);

            box = wb.createCellStyle(); setAllBorders(box); box.setFont(arial10);

            num0 = wb.createCellStyle();
            num0.cloneStyleFrom(centerBorder);
            num0.setDataFormat(fmt.getFormat("0"));
            num0.setFont(arial10);

            num1 = wb.createCellStyle();
            num1.cloneStyleFrom(centerBorder);
            num1.setDataFormat(fmt.getFormat("0.0"));
            num1.setFont(arial10);

            num2 = wb.createCellStyle();
            num2.cloneStyleFrom(centerBorder);
            num2.setDataFormat(fmt.getFormat("0.00"));
            num2.setFont(arial10);

            pmTopBottom = wb.createCellStyle();
            pmTopBottom.setAlignment(HorizontalAlignment.CENTER);
            pmTopBottom.setVerticalAlignment(VerticalAlignment.CENTER);
            pmTopBottom.setBorderTop(BorderStyle.THIN);
            pmTopBottom.setBorderBottom(BorderStyle.THIN);
            pmTopBottom.setBorderLeft(BorderStyle.NONE);
            pmTopBottom.setBorderRight(BorderStyle.NONE);
            pmTopBottom.setFont(arial10);

            // варианты без боковых граней
            num0NoRight = wb.createCellStyle(); num0NoRight.cloneStyleFrom(num0); num0NoRight.setBorderRight(BorderStyle.NONE);
            num0NoLeft  = wb.createCellStyle();  num0NoLeft.cloneStyleFrom(num0);  num0NoLeft.setBorderLeft(BorderStyle.NONE);

            num1NoRight = wb.createCellStyle(); num1NoRight.cloneStyleFrom(num1); num1NoRight.setBorderRight(BorderStyle.NONE);
            num1NoLeft  = wb.createCellStyle();  num1NoLeft.cloneStyleFrom(num1);  num1NoLeft.setBorderLeft(BorderStyle.NONE);

            num2NoRight = wb.createCellStyle(); num2NoRight.cloneStyleFrom(num2); num2NoRight.setBorderRight(BorderStyle.NONE);
            num2NoLeft  = wb.createCellStyle();  num2NoLeft.cloneStyleFrom(num2);  num2NoLeft.setBorderLeft(BorderStyle.NONE);
        }

        private static void setAllBorders(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }

    private static void setupPage(Sheet sheet) {
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        PrintSetupUtils.applyStandardMargins(sheet);
        PrintSetupUtils.applyDuplexShortEdge(sheet);
    }

    // ===== внутренние типы =====
    private record Entry(Floor floor, Space space, Room room) {}
    private static final class IntRange {
        final int min, max;
        IntRange(int min, int max) { this.min = min; this.max = max; }
        boolean isEmpty() { return max < min; }
    }
    private static final class RoomBlock {
        final Entry e;
        final int startRow;
        final double[] Gvals;
        final int[] Jvals;
        final double[] Mvals;
        final boolean isResidential, isOfficeOrPublic, isKitchen;
        RoomBlock(Entry e, int startRow, double[] g, int[] j, double[] m,
                  boolean isResidential, boolean isOfficeOrPublic, boolean isKitchen) {
            this.e = e; this.startRow = startRow; this.Gvals = g; this.Jvals = j; this.Mvals = m;
            this.isResidential = isResidential; this.isOfficeOrPublic = isOfficeOrPublic; this.isKitchen = isKitchen;
        }
    }
    private static void setupLandscapePage(Sheet sheet) {
        if (sheet == null) return;
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        PrintSetupUtils.applyStandardMargins(sheet);
        PrintSetupUtils.applyDuplexShortEdge(sheet);
    }
}
