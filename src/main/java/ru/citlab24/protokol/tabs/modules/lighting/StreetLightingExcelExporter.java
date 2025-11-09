package ru.citlab24.protokol.tabs.modules.lighting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
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

        // Ориентация/поля/умещение по ширине
        sh.getPrintSetup().setLandscape(true);
        sh.setHorizontallyCenter(true);
        sh.setAutobreaks(true);
        sh.getPrintSetup().setFitWidth((short) 1);
        sh.getPrintSetup().setFitHeight((short) 0);
        sh.setMargin(Sheet.LeftMargin,  cmToInches(1.8));
        sh.setMargin(Sheet.RightMargin, cmToInches(1.8));
        sh.setMargin(Sheet.TopMargin,   cmToInches(1.9));
        sh.setMargin(Sheet.BottomMargin,cmToInches(1.9));

        // Шрифты/стили
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

        Font f9 = wb.createFont();
        f9.setFontName("Arial");
        f9.setFontHeightInPoints((short)9);

        CellStyle header9 = wb.createCellStyle();
        header9.cloneStyleFrom(header);
        header9.setFont(f9);

        CellStyle header9Rot90 = wb.createCellStyle();
        header9Rot90.cloneStyleFrom(header9);
        header9Rot90.setRotation((short)90);

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

        CellStyle cellCenter9 = wb.createCellStyle();
        cellCenter9.cloneStyleFrom(cellCenter);
        cellCenter9.setFont(f9);

        // G/H/I стили без отдельных граней
        CellStyle gStyle = wb.createCellStyle();
        gStyle.cloneStyleFrom(cellCenter);
        gStyle.setBorderRight(BorderStyle.NONE);

        CellStyle hStyle = wb.createCellStyle();
        hStyle.cloneStyleFrom(cellCenter);
        hStyle.setBorderLeft(BorderStyle.NONE);
        hStyle.setBorderRight(BorderStyle.NONE);

        CellStyle iStyle = wb.createCellStyle();
        iStyle.cloneStyleFrom(cellCenter);
        iStyle.setBorderLeft(BorderStyle.NONE);

        CellStyle vert90Center = wb.createCellStyle();
        vert90Center.cloneStyleFrom(cellCenter);
        vert90Center.setRotation((short)90);

        // Ширины (C = 10 см); K..AD = 0,85 см
        double[] cm = {1.30, 1.03, 10.00, 2.35, 1.61, 1.16, 1.69, 0.70, 1.48, 2.51};
        for (int c = 0; c < cm.length; c++) sh.setColumnWidth(c, cmToColumnWidthUnits(cm[c]));
        for (int c = 10; c <= 29; c++)      sh.setColumnWidth(c, cmToColumnWidthUnits(0.85));

        // Высоты строк
        sh.setDefaultRowHeightInPoints((float) cmToPoints(0.45));
        setRowHeight(sh, 1, 0.45);
        setRowHeight(sh, 2, 0.11);
        setRowHeight(sh, 3, 1.19);
        setRowHeight(sh, 4, 1.53);
        setRowHeight(sh, 5, 2.46);
        setRowHeight(sh, 6, 0.45);

        // A1
        Row r1 = getOrCreateRow(sh, 0);
        Cell a1 = getOrCreateCell(r1, 0);
        a1.setCellStyle(base);
        a1.setCellValue(" 18.3 Искусственное освещение:");

        // Шапка 3–5
        merge(sh, 2, 4, 0, 0); writeHeader(sh, 2, 0, "№ п/п", header9);
        merge(sh, 2, 4, 1, 1); writeHeader(sh, 2, 1, "№ поля", header9);
        merge(sh, 2, 4, 2, 2); writeHeader(sh, 2, 2, "Наименование места\nпроведения измерений", header9);
        merge(sh, 2, 4, 3, 3); writeHeader(sh, 2, 3,
                "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м",
                header9Rot90);
        merge(sh, 2, 4, 4, 4); writeHeader(sh, 2, 4, "Вид, тип светильников", header9Rot90);
        merge(sh, 2, 4, 5, 5); writeHeader(sh, 2, 5, "Число не горящих ламп, шт.", header9Rot90);
        merge(sh, 2, 3, 6, 9); writeHeader(sh, 2, 6, "Средняя горизонтальная освещенность на уровне земли, лк", header9);
        merge(sh, 4, 4, 6, 8); writeHeader(sh, 4, 6, "измеренная ± расширенная неопределенность", header9);
        writeHeader(sh, 4, 9, "Нормируемая", header9);

        // Строка 6 — образец (центр, кегль 9)
        Row r6 = getOrCreateRow(sh, 5);
        write(sh, r6, 0, "1", cellCenter9);
        write(sh, r6, 1, "2", cellCenter9);
        write(sh, r6, 2, "3", cellCenter9);
        write(sh, r6, 3, "4", cellCenter9);
        write(sh, r6, 4, "5", cellCenter9);
        write(sh, r6, 5, "6", cellCenter9);
        merge(sh, 5, 5, 6, 8);
        write(sh, r6, 6, "7", cellCenter9);
        write(sh, r6, 9, "8", cellCenter9);

        // Новая строка 7
        Row r7 = sh.createRow(6);
        merge(sh, 6, 6, 0, 9);
        Cell t = getOrCreateCell(r7, 0);
        t.setCellStyle(cellCenter);
        t.setCellValue("Искусственная освещенность (придомовой территории и входов в здание)");

        // Подписи K7..AD7: 1..10, 1..10
        for (int c = 10; c <= 29; c++) {
            int label = ((c - 10) % 10) + 1;
            Cell cc = getOrCreateCell(r7, c);
            cc.setCellStyle(cellCenter);
            cc.setCellValue(label);
        }

        // Данные: начиная с 8-й строки (index 7)
        final int dataStartExcelRow = 8;
        int rowIndex = 7;

        for (int i = 0; i < (rows == null ? 0 : rows.size()); i++) {
            RowData rd = rows.get(i);
            int excelRow = rowIndex + 1;
            Row rr = getOrCreateRow(sh, rowIndex++);

            String num = String.valueOf(i + 1);

            // A,B
            write(sh, rr, 0, num, cellCenter);
            write(sh, rr, 1, num, cellCenter);

            // C
            write(sh, rr, 2, rd.name != null ? rd.name : "", cell);

            // D
            write(sh, rr, 3, "Г-0,0", cellCenter);

            // F
            write(sh, rr, 5, "0", cellCenter);

            // G (среднее) — EN: ROUND(SUM(K:AD)/20,1), без правой грани
            Cell cg = getOrCreateCell(rr, 6);
            cg.setCellStyle(gStyle);
            cg.setCellFormula(String.format("ROUND(SUM(K%d:AD%d)/20,1)", excelRow, excelRow));

            // H — «±», без левой/правой
            write(sh, rr, 7, "±", hStyle);

            // I — EN формула 2*SQRT( ( Σ POWER(col*0.08/3,2) ) / 20 ), без левой грани
            Cell ci = getOrCreateCell(rr, 8);
            ci.setCellStyle(iStyle);
            ci.setCellFormula(buildIFormula(excelRow));

            // J — пусто
            write(sh, rr, 9, "", cellCenter);

            // --- Инициализация K..AD нулями, чтобы формулы всегда считались ---
            for (int c = 10; c <= 29; c++) {
                Cell m = getOrCreateCell(rr, c);
                m.setCellStyle(cellCenter);
                m.setCellValue(0.0); // числовой ноль
            }

            // Размещение 4-х значений
            if (rd.leftMax   != null) getOrCreateCell(rr, 10).setCellValue(rd.leftMax.doubleValue());   // K
            if (rd.centerMin != null) getOrCreateCell(rr, 14).setCellValue(rd.centerMin.doubleValue()); // O
            if (rd.rightMax  != null) getOrCreateCell(rr, 19).setCellValue(rd.rightMax.doubleValue());  // T
            if (rd.bottomMin != null) getOrCreateCell(rr, 29).setCellValue(rd.bottomMin.doubleValue()); // AD
        }

        // E: объединённая вертикальная подпись на все строки данных
        if (rows != null && !rows.isEmpty()) {
            int lastExcelRow = dataStartExcelRow + rows.size() - 1;
            merge(sh, dataStartExcelRow - 1, lastExcelRow - 1, 4, 4);
            Row top = getOrCreateRow(sh, dataStartExcelRow - 1);
            Cell ce = getOrCreateCell(top, 4);
            ce.setCellStyle(vert90Center);
            ce.setCellValue("Общий, светодиодные");
        }

        // Область печати: только A..J
        int lastUsedRow = Math.max(sh.getLastRowNum(), rowIndex - 1);
        wb.setPrintArea(wb.getSheetIndex(sh), 0, 9, 0, lastUsedRow);
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
        CellRangeAddress region = new CellRangeAddress(r1, r2, c1, c2);
        sh.addMergedRegion(region);
        // границы по всему объединённому прямоугольнику
        RegionUtil.setBorderTop(BorderStyle.THIN,    region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN,   region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN,  region, sh);
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
    /** Merge + тонкие границы вокруг всей объединённой области. */
    private static CellRangeAddress mergeWithBorder(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress region = new CellRangeAddress(r1, r2, c1, c2);
        sh.addMergedRegion(region);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);
        return region;
    }

    private static String buildIFormula(int excelRow) {
        String[] cols = {"L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z","AA","AB","AC","AD","K"};
        StringBuilder sum = new StringBuilder();
        for (int i = 0; i < cols.length; i++) {
            if (i > 0) sum.append("+");
            sum.append("POWER(").append(cols[i]).append(excelRow).append("*0.08/3,2)");
        }
        // округление до 1 знака
        return "ROUND(2*SQRT((" + sum + ")/20),1)";
    }

}
