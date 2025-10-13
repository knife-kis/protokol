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

    private RadiationExcelExporter() {}

    /**
     * Экспорт «МЭД»:
     * - A1:I1, A2:I2, B3:I3, A4:I4 — объединено (A4 выравнено по центру).
     * - С 5-й строки — таблица с чёткими границами (тонкие чёрные), БЕЗ жирного.
     * - Секции печатаются с 9-й строки (8-я не создаётся). В строке секции пишется только её название (без «Секция:»),
     *   объединение A..I. Далее этажи этой секции:
     *      A — нумерация с 1,
     *      B — название этажа (из модели),
     *      C:E — объединено, одно число по распределению 0.10/0.11/0.12 (85%/10%/5%),
     *      F:H — объединено, одно число 0.17/0.18/0.19/0.20 (2%/13%/82%/3%), но 0.20 только если C:E > 0.10,
     *      I — пустая, но с рамкой.
     *
     * @param sectionIndex -1 = все секции подряд; >=0 = только указанная секция
     */
    public static void export(Building building, int sectionIndex, Component parent) {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sh = wb.createSheet("МЭД");

            // ===== стили =====
            Font base = wb.createFont();
            base.setFontName("Arial");
            base.setFontHeightInPoints((short)10);

            CellStyle textLeft = wb.createCellStyle();
            textLeft.setFont(base);
            textLeft.setAlignment(HorizontalAlignment.LEFT);
            textLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeft.setWrapText(true);

            CellStyle textLeftBorder = cloneWithBorders(wb, textLeft);

            CellStyle headerCenter = wb.createCellStyle();
            headerCenter.cloneStyleFrom(textLeft);
            headerCenter.setAlignment(HorizontalAlignment.CENTER);

            // без жирного
            CellStyle headerCenterBorder = cloneWithBorders(wb, headerCenter);

            CellStyle num2 = wb.createCellStyle();
            num2.setFont(base);
            num2.setAlignment(HorizontalAlignment.CENTER);
            num2.setVerticalAlignment(VerticalAlignment.CENTER);
            num2.setDataFormat(wb.createDataFormat().getFormat("0.00"));
            num2 = cloneWithBorders(wb, num2); // с рамками

            // ===== ширины колонок A..I =====
            double[] widths = {
                    4.7109375,   // A
                    51.42578125, // B
                    9.140625,    // C
                    2.85546875,  // D
                    9.0,         // E
                    6.7109375,   // F
                    2.85546875,  // G
                    9.0,         // H
                    22.28515625  // I
            };
            for (int i = 0; i < widths.length; i++) {
                sh.setColumnWidth(i, (int)Math.round(widths[i] * 256));
            }

            // ===== высоты строк (по примеру) =====
            ensureRow(sh, 4).setHeightInPoints(43.5f);  // 5-я
            ensureRow(sh, 5).setHeightInPoints(23.25f); // 6-я

            // ===== шапка =====
            // A1:I1
            merge(sh, "A1:I1");
            put(sh, 0, 0, "17. Результаты измерений ионизирующих излучений ", textLeft);

            // A2:I2
            merge(sh, "A2:I2");
            put(sh, 1, 0, "17.1. Гамма-съемка поверхности ограждающих конструкций здания с целью выявления и исключения мощных источников ", textLeft);

            // B3:I3
            merge(sh, "B3:I3");
            put(sh, 2, 1, "гамма-излучения (1 этап):", textLeft);

            // A4:I4 — центрирование, текст с 5 значениями (0.10..0.19, среднее ~0.135)
            merge(sh, "A4:I4");
            double[] gamma5 = genRow4Values(5, 0.135, 0.10, 0.19);
            put(sh, 3, 0, buildGammaLine(gamma5), headerCenter);

            // Строки 5–7: заголовки и объединения (без жирного, но с рамками)
            put(sh, 4, 0, "№ п/п", headerCenterBorder);
            put(sh, 4, 1, "Наименование места\nпроведения измерений", headerCenterBorder);
            put(sh, 4, 2, "Минимальное значение МЭД гамма-излучения, мкЗв/ч", headerCenterBorder);
            put(sh, 4, 5, "Максимальное значение МЭД гамма-излучения, мкЗв/ч", headerCenterBorder);
            put(sh, 4, 8, "Допустимый  уровень, мкЗв/ч", headerCenterBorder);

            // 'A5:A6', 'B5:B6', 'C5:E6', 'F5:H6', 'I5:I6'
            styleMerge(sh, "A5:A6", headerCenterBorder);
            styleMerge(sh, "B5:B6", headerCenterBorder);
            styleMerge(sh, "C5:E6", headerCenterBorder);
            styleMerge(sh, "F5:H6", headerCenterBorder);
            styleMerge(sh, "I5:I6", headerCenterBorder);

            // 'C7:E7', 'F7:H7' + цифры 1..5 в A7,B7,C7,F7,I7
            styleMerge(sh, "C7:E7", headerCenterBorder);
            styleMerge(sh, "F7:H7", headerCenterBorder);
            put(sh, 6, 0, 1, headerCenterBorder);
            put(sh, 6, 1, 2, headerCenterBorder);
            put(sh, 6, 2, 3, headerCenterBorder);
            put(sh, 6, 5, 4, headerCenterBorder);
            put(sh, 6, 8, 5, headerCenterBorder);

            // ===== секции и этажи =====
            // Начинаем СРАЗУ с 9-й строки (8-я не используется вообще)
            int row = 7; // 0-based → это 9-я строка в Excel
            int seq = 1;
            // Распределения
            final double[] LEFT_VALS  = {0.10, 0.11, 0.12};
            final double[] LEFT_PROB  = {0.85, 0.10, 0.05};

            final double[] RIGHT_VALS = {0.17, 0.18, 0.19, 0.20};
            final double[] RIGHT_PROB = {0.02, 0.13, 0.82, 0.03};

            DecimalFormat df2 = new DecimalFormat("0.00");

            List<Section> sections = building.getSections() != null ? building.getSections() : Collections.emptyList();
            int secStart = 0, secEnd = sections.size();
            if (sectionIndex >= 0 && sectionIndex < sections.size()) { secStart = sectionIndex; secEnd = sectionIndex + 1; }

            for (int si = secStart; si < secEnd; si++) {
                Section sec = sections.get(si);

                // этажи секции
                List<Floor> floors = new ArrayList<>();
                if (building.getFloors() != null) {
                    for (Floor f : building.getFloors()) {
                        if (f.getSectionIndex() == si) floors.add(f);
                    }
                }
                floors.sort(Comparator.comparingInt(Floor::getPosition));
// берём только этажи, где есть хотя бы одна отмеченная комната
                floors.removeIf(f -> !floorHasAnyChecked(f));
                if (floors.isEmpty()) continue; // если после фильтра ничего не осталось — пропускаем секцию

                // строка секции: объединение A..I, только название (без «Секция:»)
                String secName = (sec != null && sec.getName() != null && !sec.getName().isBlank())
                        ? sec.getName() : ("Секция " + (si + 1));
                styleMerge(sh, "A" + (row+1) + ":I" + (row+1), headerCenterBorder); // центр + рамка
                put(sh, row, 0, secName, headerCenterBorder);
                row++;

                // Этажи секции
                for (Floor f : floors) {
                    Row rr = ensureRow(sh, row);
                    Cell a = rr.getCell(0); if (a == null) a = rr.createCell(0);
                    a.setCellValue(seq++);
                    a.setCellStyle(headerCenterBorder);

                    // B — название этажа (только название из модели)
                    String floorLabel = (f.getNumber() != null && !f.getNumber().isBlank())
                            ? f.getNumber()
                            : "Этаж";
                    Cell b = rr.getCell(1); if (b == null) b = rr.createCell(1);
                    b.setCellValue(floorLabel);
                    b.setCellStyle(textLeftBorder);

                    // C:E — одно число (merge)
                    double ceVal = pickDiscrete(LEFT_VALS, LEFT_PROB);
                    styleMerge(sh, "C" + (row+1) + ":E" + (row+1), num2);
                    Cell c = rr.getCell(2); if (c == null) c = rr.createCell(2);
                    c.setCellValue(Double.parseDouble(df2.format(ceVal).replace(',', '.')));
                    c.setCellStyle(num2);

                    // F:H — одно число (merge) c запретом 0.20, если C:E <= 0.10
                    double[] rp = RIGHT_PROB.clone();
                    if (ceVal <= 0.10) {
                        rp[3] = 0.0;
                        double s = rp[0] + rp[1] + rp[2];
                        if (s > 0) { rp[0] /= s; rp[1] /= s; rp[2] /= s; }
                    }
                    double fhVal = pickDiscrete(RIGHT_VALS, rp);
                    styleMerge(sh, "F" + (row+1) + ":H" + (row+1), num2);
                    Cell fcell = rr.getCell(5); if (fcell == null) fcell = rr.createCell(5);
                    fcell.setCellValue(Double.parseDouble(df2.format(fhVal).replace(',', '.')));
                    fcell.setCellStyle(num2);

                    // I — пустая, но с рамкой
                    Cell ic = rr.getCell(8); if (ic == null) ic = rr.createCell(8);
                    ic.setCellStyle(headerCenterBorder);

                    row++;
                }
            }

            // ВАЖНО: НЕ создаём заморозку панелей → никакой «линии после 7 строки»

            // Сохранение
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить МЭД");
            chooser.setSelectedFile(new File("МЭД.xlsx"));
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

    // ===== helpers =====

    private static Row ensureRow(Sheet sh, int r0) {
        Row r = sh.getRow(r0);
        if (r == null) r = sh.createRow(r0);
        return r;
    }

    private static void put(Sheet sh, int r0, int c0, String val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = r.getCell(c0);
        if (c == null) c = r.createCell(c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void put(Sheet sh, int r0, int c0, int val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = r.getCell(c0);
        if (c == null) c = r.createCell(c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void merge(Sheet sh, String addr) {
        sh.addMergedRegion(CellRangeAddress.valueOf(addr));
    }

    /** Применить стиль на ВСЕ клетки в merged-диапазоне, плюс выполнить merge */
    private static void styleMerge(Sheet sh, String addr, CellStyle style) {
        CellRangeAddress range = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(range);
        for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
            Row row = ensureRow(sh, r);
            for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
            }
        }
    }

    private static CellStyle cloneWithBorders(Workbook wb, CellStyle src) {
        CellStyle s = wb.createCellStyle();
        s.cloneStyleFrom(src);
        s.setBorderBottom(BorderStyle.THIN);
        s.setBorderTop(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN);
        s.setBorderRight(BorderStyle.THIN);
        return s;
    }

    // ===== генерация =====
    private static boolean floorHasAnyChecked(Floor floor) {
        if (floor == null) return false;
        for (Space s : floor.getSpaces()) {
            for (Room r : s.getRooms()) {
                if (r != null && r.isSelected()) return true;
            }
        }
        return false;
    }

    /** 5 чисел 0.10–0.19 с средним близко к 0.135 */
    private static double[] genRow4Values(int n, double targetMean, double lo, double hi) {
        double mode = 0.13;
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
        DecimalFormat df = new DecimalFormat("0.00");
        return "Мощность дозы гамма-излучения на открытой местности в пяти точках составила: "
                + df.format(vals[0]).replace(',', '.') + "; "
                + df.format(vals[1]).replace(',', '.') + "; "
                + df.format(vals[2]).replace(',', '.') + "; "
                + df.format(vals[3]).replace(',', '.') + "; "
                + df.format(vals[4]).replace(',', '.') + " (мкЗв/ч)";
    }

    private static double triangular(double a, double b, double c) {
        double u = ThreadLocalRandom.current().nextDouble();
        double F = (c - a) / (b - a);
        if (u < F) return a + Math.sqrt(u * (b - a) * (c - a));
        return b - Math.sqrt((1 - u) * (b - a) * (b - c));
    }

    private static double pickDiscrete(double[] values, double[] probs) {
        double u = ThreadLocalRandom.current().nextDouble();
        double acc = 0.0;
        for (int i = 0; i < values.length; i++) {
            acc += probs[i];
            if (u <= acc) return values[i];
        }
        return values[values.length - 1];
    }

    private static double sum(double[] arr, int len) {
        double s = 0.0;
        for (int i = 0; i < len; i++) s += arr[i];
        return s;
    }

    private static double round2(double v) {
        return Math.round(v * 100.0) / 100.0;
    }
}
