package ru.citlab24.protokol.tabs.modules.lighting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/** Экспорт листа «Осв улица». */
public final class StreetLightingExcelExporter {

    private StreetLightingExcelExporter() {}

    /* ===== DTO для входных данных ===== */
    public static final class RowData {
        public final String name;
        public final Double leftMax;   // Макс слева
        public final Double centerMin; // Мин центр
        public final Double rightMax;  // Макс справа
        public final Double bottomMin; // Мин снизу

        public RowData(String name, Double leftMax, Double centerMin, Double rightMax, Double bottomMin) {
            this.name = name;
            this.leftMax = leftMax;
            this.centerMin = centerMin;
            this.rightMax = rightMax;
            this.bottomMin = bottomMin;
        }
    }

    /** Публичный экспорт с диалогом сохранения. */
    public static void export(List<RowData> rows, Component parent) throws Exception {
        try (Workbook wb = new XSSFWorkbook()) {
            appendToWorkbook(rows, wb);

            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel (Осв улица)");
            chooser.setSelectedFile(new File("Освещение_улица.xlsx"));
            if (chooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = chooser.getSelectedFile();
                if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                    file = new File(file.getAbsolutePath() + ".xlsx");
                }
                try (FileOutputStream out = new FileOutputStream(file)) {
                    wb.write(out);
                }
                JOptionPane.showMessageDialog(parent, "Файл сохранён:\n" + file.getAbsolutePath(),
                        "Готово", JOptionPane.INFORMATION_MESSAGE);
            }
        }
    }

    /** Заполнение книги согласно заданным габаритам/объединениям. */
    public static void appendToWorkbook(List<RowData> rows, Workbook wb) {
        Sheet sh = wb.createSheet("Осв улица");

        /* ===== Страница, поля, ориентация ===== */
        sh.getPrintSetup().setLandscape(true);      // горизонтально
        sh.setHorizontallyCenter(true);

        // Отступы в дюймах
        sh.setMargin(Sheet.LeftMargin,  cmToInches(1.8));
        sh.setMargin(Sheet.RightMargin, cmToInches(1.8));
        sh.setMargin(Sheet.TopMargin,   cmToInches(1.9));
        sh.setMargin(Sheet.BottomMargin,cmToInches(1.9));

        /* ===== Шрифты/стили (Arial) ===== */
        Font baseFont = wb.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short)10);

        CellStyle base = wb.createCellStyle();
        base.setFont(baseFont);

        CellStyle header = wb.createCellStyle();
        header.cloneStyleFrom(base);
        header.setWrapText(true);
        header.setAlignment(HorizontalAlignment.CENTER);
        header.setVerticalAlignment(VerticalAlignment.CENTER);
        header.setBorderTop(BorderStyle.THIN);
        header.setBorderBottom(BorderStyle.THIN);
        header.setBorderLeft(BorderStyle.THIN);
        header.setBorderRight(BorderStyle.THIN);

        CellStyle headerBold = wb.createCellStyle();
        headerBold.cloneStyleFrom(header);
        Font bold = wb.createFont();
        bold.setFontName("Arial");
        bold.setBold(true);
        bold.setFontHeightInPoints((short)10);
        headerBold.setFont(bold);

        CellStyle cell = wb.createCellStyle();
        cell.cloneStyleFrom(base);
        cell.setBorderTop(BorderStyle.THIN);
        cell.setBorderBottom(BorderStyle.THIN);
        cell.setBorderLeft(BorderStyle.THIN);
        cell.setBorderRight(BorderStyle.THIN);
        cell.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle cellCenter = wb.createCellStyle();
        cellCenter.cloneStyleFrom(cell);
        cellCenter.setAlignment(HorizontalAlignment.CENTER);

        /* ===== Колонки по ширине (см) =====
           A:1.30, B:1.03, C:11.54, D:2.35, E:1.61, F:1.16, G:1.69, H:0.70, I:1.48, J:2.51 */
        double[] cm = {1.30, 1.03, 11.54, 2.35, 1.61, 1.16, 1.69, 0.70, 1.48, 2.51};
        for (int c = 0; c < cm.length; c++) {
            sh.setColumnWidth(c, cmToColumnWidthUnits(cm[c]));
        }

        /* ===== Высоты строк (см): с 1-й — 0.45, 0.11, 1.19, 1.53, 2.46, 0.45; ниже — 0.45 ===== */
        sh.setDefaultRowHeightInPoints((float) cmToPoints(0.45));
        setRowHeight(sh, 1, 0.45);
        setRowHeight(sh, 2, 0.11);
        setRowHeight(sh, 3, 1.19);
        setRowHeight(sh, 4, 1.53);
        setRowHeight(sh, 5, 2.46);
        setRowHeight(sh, 6, 0.45);

        /* ===== A1 ===== */
        Row r1 = getOrCreateRow(sh, 0);
        Cell a1 = getOrCreateCell(r1, 0);
        a1.setCellStyle(base);
        a1.setCellValue(" 18.3 Искусственное освещение:");

        /* ===== Шапка: объединения и подписи =====
           (индексы 0-based: A=0,...,J=9; строки: 3→index2, 5→index4 и т.д.) */
        // A3:A5 — "№ п/п"
        merge(sh, 2, 4, 0, 0);
        writeHeader(sh, 2, 0, "№ п/п", headerBold);

        // B3:B5 — "№ поля"
        merge(sh, 2, 4, 1, 1);
        writeHeader(sh, 2, 1, "№ поля", headerBold);

        // C3:C5 — "Наименование места\nпроведения измерений"
        merge(sh, 2, 4, 2, 2);
        writeHeader(sh, 2, 2, "Наименование места\nпроведения измерений", headerBold);

        // D3:D5 — длинный текст
        merge(sh, 2, 4, 3, 3);
        writeHeader(sh, 2, 3,
                "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м",
                headerBold);

        // E3:E5 — "Вид, тип светильников"
        merge(sh, 2, 4, 4, 4);
        writeHeader(sh, 2, 4, "Вид, тип светильников", headerBold);

        // F3:F5 — "Число не горящих ламп, шт."
        merge(sh, 2, 4, 5, 5);
        writeHeader(sh, 2, 5, "Число не горящих ламп, шт.", headerBold);

        // G3:J4 — "Средняя горизонтальная освещенность на уровне земли, лк"
        merge(sh, 2, 3, 6, 9);
        writeHeader(sh, 2, 6, "Средняя горизонтальная освещенность на уровне земли, лк", headerBold);

        // G5:I5 — "измеренная ± расширенная неопределенность"
        merge(sh, 4, 4, 6, 8);
        writeHeader(sh, 4, 6, "измеренная ± расширенная неопределенность", headerBold);

        // J5 — "Нормируемая"
        writeHeader(sh, 4, 9, "Нормируемая", headerBold);

        /* ===== СТРОКА 6: образец пронумерованных колонок =====
           A6=1, B6=2, C6=3, D6=4, E6=5, F6=6, G6:I6=7, J6=8 */
        Row r6 = getOrCreateRow(sh, 5);
        write(sh, r6, 0, "1", cellCenter);
        write(sh, r6, 1, "2", cellCenter);
        write(sh, r6, 2, "3", cell);
        write(sh, r6, 3, "4", cellCenter);
        write(sh, r6, 4, "5", cellCenter);
        write(sh, r6, 5, "6", cellCenter);
        merge(sh, 5, 5, 6, 8);
        write(sh, r6, 6, "7", cellCenter);
        write(sh, r6, 9, "8", cellCenter);

        /* ===== Данные ниже (по строкам rows) — начинаем с 7-й строки (index 6) ===== */
        int rowIndex = 6;
        for (RowData rd : rows) {
            Row rr = getOrCreateRow(sh, rowIndex++);
            // A: № п/п — авто-нумерация
            write(sh, rr, 0, String.valueOf(rowIndex - 6), cellCenter);
            // B: № поля — пусто (пользователь заполнит позже)
            write(sh, rr, 1, "", cellCenter);
            // C: Наименование места — берём name
            write(sh, rr, 2, rd.name != null ? rd.name : "", cell);
            // D: плоскость/высота — пусто
            write(sh, rr, 3, "", cellCenter);
            // E: тип светильников — пусто
            write(sh, rr, 4, "", cellCenter);
            // F: не горят — пусто
            write(sh, rr, 5, "", cellCenter);
            // G:H:I — измеренная ± U (запишем как три ячейки: «значение», «±», «U» — пока просто значения)
            merge(sh, rr.getRowNum(), rr.getRowNum(), 6, 8);
            String val = joinNonNull(rd.leftMax, rd.centerMin, rd.rightMax, rd.bottomMin);
            write(sh, rr, 6, val, cellCenter);
            // J: нормируемая — пусто
            write(sh, rr, 9, "", cellCenter);
        }
    }

    /* ===== хелперы записи ===== */
    private static Row getOrCreateRow(Sheet sh, int idx) {
        Row r = sh.getRow(idx);
        return (r != null) ? r : sh.createRow(idx);
    }
    private static Cell getOrCreateCell(Row r, int c) {
        Cell cell = r.getCell(c);
        return (cell != null) ? cell : r.createCell(c);
    }
    private static void writeHeader(Sheet sh, int row, int col, String text, CellStyle style) {
        Row r = getOrCreateRow(sh, row);
        Cell c = getOrCreateCell(r, col);
        c.setCellStyle(style);
        c.setCellValue(text);
    }
    private static void write(Sheet sh, Row r, int col, String text, CellStyle style) {
        Cell c = getOrCreateCell(r, col);
        c.setCellStyle(style);
        c.setCellValue(text);
    }
    private static void merge(Sheet sh, int r1, int r2, int c1, int c2) {
        sh.addMergedRegion(new CellRangeAddress(r1, r2, c1, c2));
    }

    /* ===== размеры ===== */
    private static void setRowHeight(Sheet sh, int oneBasedRow, double cm) {
        Row r = getOrCreateRow(sh, oneBasedRow - 1);
        r.setHeightInPoints((float) cmToPoints(cm));
    }
    private static double cmToPoints(double cm) { return cm * 28.3464567; }        // 1 см = 28.346... pt
    private static double cmToInches(double cm) { return cm / 2.54; }              // 1 inch = 2.54 cm

    /** Приближение перевода см → units(1/256 char). */
    private static int cmToColumnWidthUnits(double cm) {
        // см → дюймы → пиксели (96 dpi) → Excel width units
        double pixels = cmToInches(cm) * 96.0;
        int widthUnits = (int) ((pixels - 5) / 7.0 * 256.0);
        return Math.max(256, widthUnits); // минимум одна «стандартная» ширина
    }

    /* ===== парсер чисел для snapshot ===== */
    public static Double tryParse(Object v) {
        if (v == null) return null;
        if (v instanceof Number n) return n.doubleValue();
        String s = v.toString().trim().replace(',', '.');
        if (s.isEmpty()) return null;
        try { return Double.parseDouble(s); } catch (Exception ignored) { return null; }
    }

    /* Склеим четыре числа в одну строку, если заданы */
    private static String joinNonNull(Double a, Double b, Double c, Double d) {
        StringBuilder sb = new StringBuilder();
        if (a != null) sb.append(a);
        if (b != null) { if (!sb.isEmpty()) sb.append(" ; "); sb.append(b); }
        if (c != null) { if (!sb.isEmpty()) sb.append(" ; "); sb.append(c); }
        if (d != null) { if (!sb.isEmpty()) sb.append(" ; "); sb.append(d); }
        return sb.isEmpty() ? "" : sb.toString();
    }
}
