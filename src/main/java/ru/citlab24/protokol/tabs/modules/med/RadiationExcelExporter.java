package ru.citlab24.protokol.tabs.modules.med;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;

public final class RadiationExcelExporter {
    private static final String LIMIT_TEXT =
            "Превышение мощности дозы, измеренной на открытой местности, не более чем на 0,3 мкЗв/ч";


    private RadiationExcelExporter() {}

    /** Вызов из вкладки: экспорт обеих листов — «МЭД» и «МЭД (2)». */
    public static void export(Building building, int sectionIndex, Component parent) {
        try (Workbook wb = new XSSFWorkbook()) {

            // ===== общие стили =====
            Styles S = new Styles(wb);

            // ===== 1) Лист «МЭД» (как согласовали) =====
            double[] gamma5 = buildSheetMED(wb, building, sectionIndex, S);

            // ===== 2) Лист «МЭД (2)» (новый по твоему примеру) =====
            buildSheetMED2(wb, building, sectionIndex, S, gamma5);

// ===== 3) Лист «ЭРОА радона» =====
            buildSheetRadon(wb, building, sectionIndex, S);

// ===== сохранение =====
            JFileChooser chooser = new JFileChooser();

            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Ионизирующее_излучение.xlsx"));
            if (chooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = chooser.getSelectedFile();
                if (!file.getName().toLowerCase().endsWith(".xlsx"))
                    file = new File(file.getAbsolutePath() + ".xlsx");
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

    /* ============================ Лист 1: «МЭД» ============================ */

    private static double[] buildSheetMED(Workbook wb, Building building, int sectionIndex, Styles S) {
        Sheet sh = wb.createSheet("МЭД");

        // ширины A..I
        setColWidths(sh, new double[]{4.71, 51.43, 9.14, 2.86, 9.0, 6.71, 2.86, 9.0, 22.29});

        // высоты (5-я, 6-я)
        ensureRow(sh, 4).setHeightInPoints(43.5f);
        ensureRow(sh, 5).setHeightInPoints(23.25f);

        // A1:I1
        merge(sh, "A1:I1"); put(sh, 0, 0, "17. Результаты измерений ионизирующих излучений ", S.textLeft);

        // A2:I2
        merge(sh, "A2:I2"); put(sh, 1, 0, "17.1. Гамма-съемка поверхности ограждающих конструкций здания с целью выявления и исключения мощных источников ", S.textLeft);

        // B3:I3
        merge(sh, "B3:I3"); put(sh, 2, 1, "гамма-излучения (1 этап):", S.textLeft);

        // A4:I4 — 5 случайных значений 0.10..0.19 (ср≈0.135)
        merge(sh, "A4:I4");
        double[] gamma5 = genRow4Values(5, 0.135, 0.10, 0.19);
        put(sh, 3, 0, buildGammaLine(gamma5), S.headerCenter);

        // заголовки 5–7 (без жирного)
        put(sh, 4, 0, "№ п/п", S.headerCenterBorder);
        put(sh, 4, 1, "Наименование места\nпроведения измерений", S.headerCenterBorder);
        put(sh, 4, 2, "Минимальное значение МЭД гамма-излучения, мкЗв/ч", S.headerCenterBorder);
        put(sh, 4, 5, "Максимальное значение МЭД гамма-излучения, мкЗв/ч", S.headerCenterBorder);
        put(sh, 4, 8, "Допустимый  уровень, мкЗв/ч", S.headerCenterBorder);

        styleMerge(sh, "A5:A6", S.headerCenterBorder);
        styleMerge(sh, "B5:B6", S.headerCenterBorder);
        styleMerge(sh, "C5:E6", S.headerCenterBorder);
        styleMerge(sh, "F5:H6", S.headerCenterBorder);
        styleMerge(sh, "I5:I6", S.headerCenterBorder);

        styleMerge(sh, "C7:E7", S.headerCenterBorder);
        styleMerge(sh, "F7:H7", S.headerCenterBorder);
        put(sh, 6, 0, 1, S.headerCenterBorder);
        put(sh, 6, 1, 2, S.headerCenterBorder);
        put(sh, 6, 2, 3, S.headerCenterBorder);
        put(sh, 6, 5, 4, S.headerCenterBorder);
        put(sh, 6, 8, 5, S.headerCenterBorder);

        // секции с 8-й строки
        int row = 7;
        int seq = 1; // сквозная нумерация

        List<Section> sections = building.getSections() != null ? building.getSections() : Collections.emptyList();
        int secStart = 0, secEnd = sections.size();
        if (sectionIndex >= 0 && sectionIndex < sections.size()) { secStart = sectionIndex; secEnd = sectionIndex + 1; }

        final double[] LEFT_VALS  = {0.10, 0.11, 0.12};
        final double[] LEFT_PROB  = {0.85, 0.10, 0.05};
        final double[] RIGHT_VALS = {0.17, 0.18, 0.19, 0.20};
        final double[] RIGHT_PROB = {0.02, 0.13, 0.82, 0.03};
        DecimalFormat df2 = new DecimalFormat("0.00");

        for (int si = secStart; si < secEnd; si++) {
            Section sec = sections.get(si);

            List<Floor> floors = floorsOfSection(building, si);
            // только этажи, где есть хотя бы одна отмеченная комната
            floors.removeIf(f -> !floorHasAnyChecked(f));
            if (floors.isEmpty()) continue;

            boolean printSectionHeader = sections.size() > 1;
            if (printSectionHeader) {
                String secName = (sec != null && notBlank(sec.getName())) ? sec.getName() : ("Секция " + (si + 1));
                styleMerge(sh, "A" + (row+1) + ":I" + (row+1), S.headerCenterBorder);
                put(sh, row, 0, secName, S.headerCenterBorder);
                row++;
            }

            // запомним диапазон строк этажей этой секции, чтобы затем объединить I
            int dataStart = row;                // первая строка этажей
            int dataEnd   = row - 1;            // будет обновлён в цикле

            for (Floor f : floors) {
                Row rr = ensureRow(sh, row);

                // A — номер (сквозной)
                Cell a = cell(rr, 0); a.setCellValue(seq++); a.setCellStyle(S.headerCenterBorder);

                // B — название этажа
                String floorLabel = notBlank(f.getNumber()) ? f.getNumber() : "Этаж";
                Cell b = cell(rr, 1); b.setCellValue(floorLabel); b.setCellStyle(S.textLeftBorder);

                // C:E — минимальное
                double ceVal = pickDiscrete(LEFT_VALS, LEFT_PROB);
                styleMerge(sh, "C" + (row+1) + ":E" + (row+1), S.num2);
                Cell c = cell(rr, 2); c.setCellValue(parse2(df2, ceVal)); c.setCellStyle(S.num2);

                // F:H — максимальное (с условием по 0.20)
                double[] rp = RIGHT_PROB.clone();
                if (ceVal <= 0.10) {
                    rp[3] = 0.0; double s = rp[0]+rp[1]+rp[2]; if (s>0){rp[0]/=s; rp[1]/=s; rp[2]/=s;}
                }
                double fhVal = pickDiscrete(RIGHT_VALS, rp);
                styleMerge(sh, "F" + (row+1) + ":H" + (row+1), S.num2);
                Cell fcell = cell(rr, 5); fcell.setCellValue(parse2(df2, fhVal)); fcell.setCellStyle(S.num2);

                // I — пока только рамка (значение поставим после merge секции)
                Cell ic = cell(rr, 8); ic.setCellStyle(S.headerCenterBorder);

                dataEnd = row;
                row++;
            }

            // объединяем I по всем строкам этажей этой секции и вписываем текст
            if (dataEnd >= dataStart) {
                String rng = "I" + (dataStart+1) + ":I" + (dataEnd+1);
                styleMerge(sh, rng, S.headerCenterBorder);
                put(sh, dataStart, 8, LIMIT_TEXT, S.headerCenterBorder);
            }
        }

        return gamma5;
    }


    /* ============================ Лист 2: «МЭД (2)» ============================ */

    private static void buildSheetMED2(Workbook wb, Building building, int sectionIndex, Styles S, double[] gamma5) {
        final String LIMIT_TEXT = "Превышение мощности дозы, измеренной на открытой местности, не более чем на 0,3 мкЗв/ч";

        Sheet sh = wb.createSheet("МЭД (2)");

        // ширины A..F
        setColWidths(sh, new double[]{4.71, 60.0, 9.14, 3.0, 9.14, 36.0});

        ensureRow(sh, 2).setHeightInPoints(29.25f);

        // ===== шапка (первые 5 строк фиксированы) =====
        merge(sh, "A1:F1");
        put(sh, 0, 0, "17.2. Мощность дозы гамма-излучений (2 этап):", S.textLeft);

        merge(sh, "A2:F2");
        // те же значения, что на листе «МЭД»
        put(sh, 1, 0, "Мощность дозы гамма-излучения на открытой местности в пяти точках составила: "
                + joinGamma(gamma5) + " (мкЗв/ч)", S.textLeft);

        // заголовки таблицы на 3–5 строках
        // A3:A4, B3:B4, C3:E4, F3:F4
        put(sh, 2, 0, "№ п/п", S.headerCenterBorder);
        put(sh, 2, 1, "Наименование места\nпроведения измерений", S.headerCenterBorder);
        put(sh, 2, 2, "Мощность дозы гамма-\nизлучения ± ΔН, мкЗв/ч", S.headerCenterBorder);
        put(sh, 2, 5, "Допустимый  уровень, мкЗв/ч", S.headerCenterBorder);

        styleMerge(sh, "A3:A4", S.headerCenterBorder);
        styleMerge(sh, "B3:B4", S.headerCenterBorder);
        styleMerge(sh, "C3:E4", S.headerCenterBorder);
        styleMerge(sh, "F3:F4", S.headerCenterBorder);

        // строка 5 — нумерация 1..4 как в примере (C..E объединены)
        styleMerge(sh, "C5:E5", S.headerCenterBorder);
        put(sh, 4, 0, 1, S.headerCenterBorder);
        put(sh, 4, 1, 2, S.headerCenterBorder);
        put(sh, 4, 2, 3, S.headerCenterBorder);
        put(sh, 4, 5, 4, S.headerCenterBorder);

        // ===== данные с 6-й строки =====
        int row = 4;     // 0-based → 6-я строка
        int seq = 1;     // сквозная нумерация помещений

        List<Section> sections = building.getSections() != null ? building.getSections() : Collections.emptyList();
        int secStart = 0, secEnd = sections.size();
        if (sectionIndex >= 0 && sectionIndex < sections.size()) { secStart = sectionIndex; secEnd = sectionIndex + 1; }

        DecimalFormat df2 = new DecimalFormat("0.00");

        for (int si = secStart; si < secEnd; si++) {
            Section sec = sections.get(si);
            String secName = (sec != null && notBlank(sec.getName())) ? sec.getName() : ("Секция " + (si + 1));

            List<Floor> floors = floorsOfSection(building, si);
            // оставляем только этажи, где есть хотя бы одна отмеченная комната
            floors.removeIf(f -> !floorHasAnyChecked(f));
            if (floors.isEmpty()) continue;

            for (Floor f : floors) {
                // строка с «[Секция, ]Этаж» (секцию не пишем, если она единственная)
                row++;
                String floorTitle = notBlank(f.getNumber()) ? f.getNumber() : "Этаж";
                String header = (sections.size() > 1 ? (secName + ", ") : "") + floorTitle;
                styleMerge(sh, "A" + (row+1) + ":F" + (row+1), S.headerCenterBorder);
                put(sh, row, 0, header, S.headerCenterBorder);


                // комнаты этого этажа (только отмеченные)
                List<RoomEntry> entries = checkedRoomEntriesOnFloor(f);

                // --- границы блока по F: первая и последняя строка данных
                int dataStart = row + 1;            // первая строка с комнатами (0-based)
                int dataEnd   = row;                 // временно (если нет комнат)

                for (RoomEntry re : entries) {
                    row++;
                    Row rr = ensureRow(sh, row);

                    // A — сквозная нумерация
                    Cell a = cell(rr, 0); a.setCellValue(seq++); a.setCellStyle(S.headerCenterBorder);

                    // B — "<помещение>, <комната>"
                    String roomLabel = safeName(re.room.getName());
                    String bText = isPublicSpace(re.space) ? roomLabel : (spaceDisplayName(re.space) + ", " + roomLabel);
                    Cell b = cell(rr, 1); b.setCellValue(bText); b.setCellStyle(S.textLeftBorder);


                    // C — значение (ср. ≈ 0,135)
                    double cVal = sampleMEDValue();
                    Cell c = cell(rr, 2); c.setCellValue(parse2(df2, cVal)); c.setCellStyle(S.num2NoRight);

                    // D — "±" с бордерами только сверху/снизу
                    Cell d = cell(rr, 3); d.setCellValue("±"); d.setCellStyle(S.plusMinusTB);

                    // E — ΔH
                    double delta = ((15.0 + (4.0 / cVal * 0.01)) * cVal / 100.0);
                    Cell e = cell(rr, 4); e.setCellValue(parse2(df2, delta)); e.setCellStyle(S.num2NoLeft);

                    // F — пока просто стиль (значение поставим после merge)
                    Cell fcell = cell(rr, 5); fcell.setCellStyle(S.headerCenterBorder);

                    dataEnd = row; // последняя строка данных
                }

                // NEW: объединяем F по всему блоку комнат и пишем текст
                if (!entries.isEmpty()) {
                    String rng = "F" + (dataStart+1) + ":F" + (dataEnd+1); // A1-индексация
                    styleMerge(sh, rng, S.headerCenterBorder);
                    put(sh, dataStart, 5, LIMIT_TEXT, S.headerCenterBorder); // значение в верхнюю ячейку merged-области
                }
            }
        }
    }
    /* ============================ Лист 3: «ЭРОА радона» ============================ */

    private static void buildSheetRadon(Workbook wb, Building building, int sectionIndex, Styles S) {
        Sheet sh = wb.createSheet("ЭРОА радона");

        // ширины A..G
        setColWidths(sh, new double[]{4.71, 60.0, 9.5, 6.0, 11.5, 14.0, 18.0});
        setColWidthPx(sh, 1, 333); // B
        setColWidthPx(sh, 2, 104); // C
        setColWidthPx(sh, 3, 104); // D
        setColWidthPx(sh, 4, 104); // E
        // ===== шапка (строки 1–5) как в образце ЭРОА.xlsx =====

// 1-я строка — общий заголовок
        merge(sh, "A1:G1");
        put(sh, 0, 0, "17.3. ЭРОА радона, ЭРОА торона, среднегодовое значение ЭРОА изотопов радона:", S.textLeft);

// 2-я строка — пустая техническая (для отступа/границ)
        Row r2 = ensureRow(sh, 1);
        r2.setHeightInPoints(3f);
        for (int c = 0; c <= 6; c++) {
            cell(r2, c).setCellStyle(S.bottomOnly);
        }

// 3-я строка — внешний уровень
        styleMerge(sh, "A3:A4", S.headerCenterBorder);
        put(sh, 2, 0, "№ п/п", S.headerCenterBorder);

        styleMerge(sh, "B3:B4", S.headerCenterBorder);
        put(sh, 2, 1, "Наименование места\nпроведения измерений", S.headerCenterBorder);

// «Результаты измерений, Бк/м³» на C3:F3
        styleMerge(sh, "C3:F3", S.headerCenterBorder);
        put(sh, 2, 2, "Результаты измерений, Бк/м³", S.headerCenterBorder);

// Правый столбец
        styleMerge(sh, "G3:G4", S.headerCenterBorder);
        put(sh, 2, 6, "Допустимый уровень, Бк/м³", S.headerCenterBorder);

// 4-я строка — внутренние подписи блока результатов
        styleMerge(sh, "C4:D4", S.headerCenterBorder);
        put(sh, 3, 2, "Измеренное значение ЭРОА радона (ЭРОА торона)", S.headerCenterBorder);

        put(sh, 3, 4, "Среднегодовое значение ЭРОА\nизотопов радона, Бк/м³", S.headerCenterBorder);
        put(sh, 3, 5, "Суммарная\nнеопределённость", S.headerCenterBorder);

// 5-я строка — номера столбцов (как в образце) + объединение C5:D5
        put(sh, 4, 0, 1, S.headerCenterBorder);
        put(sh, 4, 1, 2, S.headerCenterBorder);
        styleMerge(sh, "C5:D5", S.headerCenterBorder);
        put(sh, 4, 2, 3, S.headerCenterBorder);
        put(sh, 4, 4, 4, S.headerCenterBorder);
        put(sh, 4, 5, 5, S.headerCenterBorder);
        put(sh, 4, 6, 7, S.headerCenterBorder);


        // данные с 6-й строки
        int row = 4;     // 0-based → 6-я
        int seq = 1;     // № п/п

        List<Section> sections = building.getSections() != null ? building.getSections() : Collections.emptyList();
        int secStart = 0, secEnd = sections.size();
        if (sectionIndex >= 0 && sectionIndex < sections.size()) { secStart = sectionIndex; secEnd = sectionIndex + 1; }

        // сезонный коэффициент: май–сентябрь = 1.3, иначе 1.0
        Calendar cal = Calendar.getInstance();
        int m = cal.get(Calendar.MONTH) + 1; // 1..12
        double seasonK = (m >= 5 && m <= 9) ? 1.3 : 1.0;
        String seasonKStr = (seasonK == 1.3) ? "1.3" : "1.0";

        ThreadLocalRandom rnd = ThreadLocalRandom.current();

        for (int si = secStart; si < secEnd; si++) {
            Section sec = sections.get(si);
            String secName = (sec != null && notBlank(sec.getName())) ? sec.getName() : ("Секция " + (si + 1));

            List<Floor> floors = floorsOfSection(building, si);
            floors.removeIf(f -> !floorHasAnyChecked(f));
            if (floors.isEmpty()) continue;

            for (Floor f : floors) {
                // строка "Секция, Этаж"
                row++;
                String floorTitle = notBlank(f.getNumber()) ? f.getNumber() : "Этаж";
                String header = (sections.size() > 1 ? (secName + ", ") : "") + floorTitle;
                styleMerge(sh, "A" + (row+1) + ":G" + (row+1), S.headerCenterBorder);
                put(sh, row, 0, header, S.headerCenterBorder);


                // элементы (офисы/общественные — ВСЕ отмеченные; квартиры — по 1 комнате)
                List<RadonEntry> entries = radonEntriesOnFloor(f);

                int dataStart = row + 1;
                int dataEnd   = row;

                for (RadonEntry re : entries) {
                    row++;
                    Row rr = ensureRow(sh, row);

                    // A — № п/п
                    Cell a = cell(rr, 0); a.setCellValue(seq++); a.setCellStyle(S.headerCenterBorder);

                    // B — "<помещение>, комнаты через запятую"
                    String roomName = re.rooms.isEmpty() ? "" : safeName(re.rooms.get(0).getName());
                    String bText = isPublicSpace(re.space)
                            ? roomName
                            : spaceDisplayName(re.space) + (roomName.isBlank() ? "" : ", " + roomName);
                    Cell b = cell(rr, 1); b.setCellValue(bText); b.setCellStyle(S.textLeftBorder);

                    // C — случайное целое 17..40
                    int cVal = rnd.nextInt(17, 41);
                    Cell c = cell(rr, 2); c.setCellValue(cVal); c.setCellStyle(S.headerCenterBorder);

                    // D — всегда 1
                    Cell d = cell(rr, 3); d.setCellValue(1); d.setCellStyle(S.headerCenterBorder);

                    // E — =(C+4.6*D)*K  (десятичные ТОЧКИ для формул Excel)
                    String baseE = String.format(java.util.Locale.US, "(C%d+4.6*D%d)*%s", row+1, row+1, seasonKStr);
                    Cell e = cell(rr, 4); e.setCellFormula("ROUND(" + baseE + ",0)"); e.setCellStyle(S.num0);

                    String baseF = String.format(java.util.Locale.US,
                            "SQRT(POWER(C%d*0.3*(2/POWER(3,0.5))*%s,2)+21.16*POWER(D%d*0.3*(2/POWER(3,0.5))*%s,2))",
                            row+1, seasonKStr, row+1, seasonKStr);
                    Cell fcell = cell(rr, 5); fcell.setCellFormula("ROUND(" + baseF + ",0)"); fcell.setCellStyle(S.num0);


                    // G — значение заполним после merge
                    Cell g = cell(rr, 6); g.setCellStyle(S.headerCenterBorder);

                    dataEnd = row;
                }

                // G объединяем на блок комнат этажа и ставим "10"
                if (!entries.isEmpty()) {
                    String rng = "G" + (dataStart+1) + ":G" + (dataEnd+1);
                    styleMerge(sh, rng, S.headerCenterBorder);
                    put(sh, dataStart, 6, 100, S.headerCenterBorder);
                }
            }
        }
    }

    private static String buildRoomsList(List<Room> rooms) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < rooms.size(); i++) {
            String n = safeName(rooms.get(i).getName());
            if (n != null && !n.isBlank()) {
                if (sb.length() > 0) sb.append(", ");
                sb.append(n);
            }
        }
        return sb.toString();
    }

    private static class RadonEntry {
        final Space space;
        final List<Room> rooms;
        RadonEntry(Space s, List<Room> r) { this.space = s; this.rooms = r; }
    }

    // Офисы/общественные — все отмеченные комнаты;
// Квартиры — если отмечено >=1, берём случайно одну.
    private static List<RadonEntry> radonEntriesOnFloor(Floor floor) {
        List<RadonEntry> res = new ArrayList<>();
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        for (Space s : floor.getSpaces()) {
            List<Room> selected = new ArrayList<>();
            for (Room r : s.getRooms()) {
                if (r != null && r.isRadiationSelected()) selected.add(r);
            }
            if (selected.isEmpty()) continue;

            Space.SpaceType tp = s.getType();
            if (tp == Space.SpaceType.APARTMENT) {
                Room r = selected.get(selected.size() == 1 ? 0 : rnd.nextInt(selected.size()));
                res.add(new RadonEntry(s, Collections.singletonList(r)));
            } else {
                for (Room r : selected) {
                    res.add(new RadonEntry(s, Collections.singletonList(r)));
                }
            }
        }
        return res;
    }

    /* ============================ Вспомогательные ============================ */

    private static class Styles {
        final CellStyle textLeft, textLeftBorder, headerCenter, headerCenterBorder, num2;
        final CellStyle plusMinusTB;
        final CellStyle num2NoRight;
        final CellStyle num2NoLeft;
        final CellStyle num0;        // целые числа
        final CellStyle bottomOnly;

        Styles(Workbook wb) {
            Font base = wb.createFont();
            base.setFontName("Arial");
            base.setFontHeightInPoints((short)10);

            textLeft = wb.createCellStyle();
            textLeft.setFont(base);
            textLeft.setAlignment(HorizontalAlignment.LEFT);
            textLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeft.setWrapText(true);

            textLeftBorder = cloneWithBorders(wb, textLeft);

            headerCenter = wb.createCellStyle();
            headerCenter.cloneStyleFrom(textLeft);
            headerCenter.setAlignment(HorizontalAlignment.CENTER);

            headerCenterBorder = cloneWithBorders(wb, headerCenter);

            num2 = wb.createCellStyle();
            num2.setFont(base);
            num2.setAlignment(HorizontalAlignment.CENTER);
            num2.setVerticalAlignment(VerticalAlignment.CENTER);
            num2.setDataFormat(wb.createDataFormat().getFormat("0.00"));
            cloneIntoBorders(num2);

            // ← стиль для столбца D (±): только верх/низ
            plusMinusTB = wb.createCellStyle();
            plusMinusTB.cloneStyleFrom(headerCenter);     // выравнивание по центру
            plusMinusTB.setBorderTop(BorderStyle.THIN);
            plusMinusTB.setBorderBottom(BorderStyle.THIN);
            plusMinusTB.setBorderLeft(BorderStyle.NONE);
            plusMinusTB.setBorderRight(BorderStyle.NONE);

            num2NoRight = wb.createCellStyle();
            num2NoRight.cloneStyleFrom(num2);
            num2NoRight.setBorderRight(BorderStyle.NONE); // убираем правую

            num2NoLeft = wb.createCellStyle();
            num2NoLeft.cloneStyleFrom(num2);
            num2NoLeft.setBorderLeft(BorderStyle.NONE);

// формат целых чисел
            num0 = wb.createCellStyle();
            num0.setFont(base);
            num0.setAlignment(HorizontalAlignment.CENTER);
            num0.setVerticalAlignment(VerticalAlignment.CENTER);
            num0.setDataFormat(wb.createDataFormat().getFormat("0"));
            cloneIntoBorders(num0);

// только нижняя граница (для строки 2 листа ЭРОА)
            bottomOnly = wb.createCellStyle();
            bottomOnly.setBorderBottom(BorderStyle.THIN);
            bottomOnly.setBorderTop(BorderStyle.NONE);
            bottomOnly.setBorderLeft(BorderStyle.NONE);
            bottomOnly.setBorderRight(BorderStyle.NONE);
            bottomOnly.setAlignment(HorizontalAlignment.CENTER);
            bottomOnly.setVerticalAlignment(VerticalAlignment.CENTER);
        }
    }

    private static void setColWidths(Sheet sh, double[] widths) {
        for (int i = 0; i < widths.length; i++) {
            sh.setColumnWidth(i, (int)Math.round(widths[i] * 256));
        }
    }

    private static Row ensureRow(Sheet sh, int r0) {
        Row r = sh.getRow(r0);
        if (r == null) r = sh.createRow(r0);
        return r;
    }
    private static Cell cell(Row r, int c0) {
        Cell c = r.getCell(c0);
        if (c == null) c = r.createCell(c0);
        return c;
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

    private static void merge(Sheet sh, String addr) {
        CellRangeAddress want = CellRangeAddress.valueOf(addr);
        for (int i = 0; i < sh.getNumMergedRegions(); i++) {
            CellRangeAddress r = sh.getMergedRegion(i);
            if (r.getFirstRow()    == want.getFirstRow() &&
                    r.getLastRow()     == want.getLastRow()  &&
                    r.getFirstColumn() == want.getFirstColumn() &&
                    r.getLastColumn()  == want.getLastColumn()) {
                return; // уже объединено — выходим
            }
        }
        sh.addMergedRegion(want);
    }
    private static void styleMerge(Sheet sh, String addr, CellStyle style) {
        CellRangeAddress range = CellRangeAddress.valueOf(addr);

        // если диапазон состоит из одной ячейки — не выполняем merge,
        // только применяем стиль
        boolean singleCell = range.getFirstRow() == range.getLastRow()
                && range.getFirstColumn() == range.getLastColumn();
        if (!singleCell) {
            merge(sh, addr); // безопасно: не добавит дубликат
        }

        for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
            Row row = ensureRow(sh, r);
            for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                Cell cell = cell(row, c);
                if (style != null) cell.setCellStyle(style);
            }
        }
    }

    private static CellStyle cloneWithBorders(Workbook wb, CellStyle src) {
        CellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(src);
        cloneIntoBorders(s);
        return s;
    }
    private static void cloneIntoBorders(CellStyle s) {
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
    }

    private static boolean notBlank(String s) { return s != null && !s.isBlank(); }
    private static String safeName(String s) { return s == null ? "" : s; }

    private static List<Floor> floorsOfSection(Building b, int secIdx) {
        List<Floor> out = new ArrayList<>();
        if (b.getFloors() != null) {
            for (Floor f : b.getFloors()) if (f.getSectionIndex() == secIdx) out.add(f);
        }
        out.sort(Comparator.comparingInt(Floor::getPosition));
        return out;
    }

    /** Есть ли на этаже хотя бы одна отмеченная комната */
    private static boolean floorHasAnyChecked(Floor floor) {
        if (floor == null) return false;
        for (Space s : floor.getSpaces()) {
            for (Room r : s.getRooms()) {
                if (r != null && r.isRadiationSelected()) return true;
            }
        }
        return false;
    }

    /** Список отмеченных комнат этого этажа (по всем помещениям) */
    private static List<Room> checkedRoomsOnFloor(Floor floor) {
        List<Room> res = new ArrayList<>();
        for (Space s : floor.getSpaces()) {
            for (Room r : s.getRooms()) if (r != null && r.isSelected()) res.add(r);
        }
        return res;
    }

    /** 5 чисел 0.10–0.19 с средним близко к 0.135 */
    private static double[] genRow4Values(int n, double targetMean, double lo, double hi) {
        double mode = 0.13; // пиковая область ~0.13
        double[] vals = new double[n];

        for (int attempt = 0; attempt < 500; attempt++) {
            double sum = 0.0;
            for (int i = 0; i < n; i++) {
                vals[i] = round2(triangular(lo, hi, mode));
                sum += vals[i];
            }
            double m = sum / n;
            if (m >= 0.132 && m <= 0.138) return vals;
        }
        for (int i = 0; i < n - 1; i++) vals[i] = round2(triangular(lo, hi, mode));
        double need = targetMean * n - sum(vals, n - 1);
        need = Math.max(lo, Math.min(hi, need));
        vals[n - 1] = round2(need);
        return vals;
    }

    private static String buildGammaLine(double[] vals) {
        return "Мощность дозы гамма-излучения на открытой местности в пяти точках составила: "
                + joinGamma(vals) + " (мкЗв/ч)";
    }
    private static String joinGamma(double[] v) {
        DecimalFormat df = new DecimalFormat("0.00");
        return df.format(v[0]).replace(',', '.') + "; "
                + df.format(v[1]).replace(',', '.') + "; "
                + df.format(v[2]).replace(',', '.') + "; "
                + df.format(v[3]).replace(',', '.') + "; "
                + df.format(v[4]).replace(',', '.');
    }

    private static double triangular(double a, double b, double c) {
        double u = ThreadLocalRandom.current().nextDouble();
        double F = (c - a) / (b - a);
        if (u < F) return a + Math.sqrt(u * (b - a) * (c - a));
        return b - Math.sqrt((1 - u) * (b - a) * (b - c));
    }
    private static double round2(double v) { return Math.round(v * 100.0) / 100.0; }
    private static double sum(double[] arr, int len) { double s=0; for(int i=0;i<len;i++) s+=arr[i]; return s; }
    private static double pickDiscrete(double[] values, double[] probs) {
        double u = ThreadLocalRandom.current().nextDouble(), acc = 0.0;
        for (int i = 0; i < values.length; i++) { acc += probs[i]; if (u <= acc) return values[i]; }
        return values[values.length - 1];
    }
    private static double parse2(DecimalFormat df, double v) {
        return Double.parseDouble(df.format(v).replace(',', '.'));
    }

    /** Сэмпл для C (0.10..0.20 со средним ~0.135) */
    private static double sampleMEDValue() {
        double[] values = {0.10, 0.11, 0.12, 0.13, 0.14, 0.15, 0.16, 0.17, 0.18, 0.19, 0.20};
        double[] probs  = {0.015,0.080,0.160,0.300,0.240,0.100,0.045,0.030,0.015,0.010,0.005};
        // Проверка: сумма = 1.000, матожидание = 0.135
        return pickDiscrete(values, probs);
    }
    private static class RoomEntry {
        final Space space;
        final Room room;
        RoomEntry(Space s, Room r) { this.space = s; this.room = r; }
    }

    private static List<RoomEntry> checkedRoomEntriesOnFloor(Floor floor) {
        List<RoomEntry> res = new ArrayList<>();
        for (Space s : floor.getSpaces()) {
            for (Room r : s.getRooms()) {
                if (r != null && r.isRadiationSelected()) {
                    res.add(new RoomEntry(s, r));
                }
            }
        }
        return res;
    }


    private static String spaceDisplayName(Space s) {
        if (s == null) return "";
        String id = s.getIdentifier();
        if (id != null && !id.isBlank()) return id;
        // если в модели есть getName() — можно подставить:
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
    private static void setColWidthPx(Sheet sh, int col, int px) {
        int width = (int) Math.round((px - 5) / 7.0 * 256); // прибл. формула Excel
        if (width < 0) width = 0;
        sh.setColumnWidth(col, width);
    }
    private static boolean isPublicSpace(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String txt = (t.name() + " " + String.valueOf(t)).toLowerCase(java.util.Locale.ROOT);
        return txt.contains("public") || txt.contains("обще");
    }

}
