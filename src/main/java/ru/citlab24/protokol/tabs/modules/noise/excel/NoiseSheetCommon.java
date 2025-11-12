package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind;

import java.io.File;
import java.util.Map;

public class NoiseSheetCommon {

    private NoiseSheetCommon() {}

    /* ===== Общие помощники по странице/колонкам/печати ===== */

    public static void setupPage(Sheet sh) {
        PrintSetup ps = sh.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);
        sh.setMargin(Sheet.LeftMargin,  1.80 / 2.54);
        sh.setMargin(Sheet.RightMargin, 1.48 / 2.54);
        sh.setAutobreaks(true);
        sh.setFitToPage(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
    }

    public static void setupColumns(Sheet sh) {
        java.util.function.BiConsumer<Integer, Double> setCm = (col, cm) -> {
            double chars = cm * 5.4;
            sh.setColumnWidth(col, (int) Math.round(chars * 256));
        };
        setCm.accept(0, 0.82);  // A
        setCm.accept(1, 0.95);  // B
        setCm.accept(2, 4.55);  // C
        setCm.accept(3, 3.70);  // D
        for (int c = 4; c <= 8; c++) setCm.accept(c, 0.69);  // E..I
        for (int c = 9; c <= 17; c++) setCm.accept(c, 0.95); // J..R
        setCm.accept(18, 0.90); // S
        setCm.accept(19, 0.71); // T
        setCm.accept(20, 0.32); // U
        setCm.accept(21, 0.58); // V
        setCm.accept(22, 0.71); // W
        setCm.accept(23, 0.32); // X
        setCm.accept(24, 0.64); // Y
    }

    /** Сжать ширины A..Y в 1 страницу по ширине. */
    public static void shrinkColumnsToOnePrintedPage(Sheet sh) {
        final int FIRST_COL = 0, LAST_COL = 24;
        long sum = 0;
        for (int c = FIRST_COL; c <= LAST_COL; c++) sum += sh.getColumnWidth(c);
        if (sum <= 0) return;

        PrintSetup ps = sh.getPrintSetup();
        double pageWidthCm = ps.getLandscape() ? 29.7 : 21.0;
        double leftCm = sh.getMargin(Sheet.LeftMargin) * 2.54;
        double rightCm = sh.getMargin(Sheet.RightMargin) * 2.54;
        double printableCm = Math.max(0.1, pageWidthCm - (leftCm + rightCm));
        long targetUnits = (long) Math.floor((printableCm * 5.4) * 256.0);

        if (sum <= targetUnits) return;
        double k = (double) targetUnits / (double) sum;
        final int MIN_UNITS = 1 * 256;
        for (int c = FIRST_COL; c <= LAST_COL; c++) {
            int w = sh.getColumnWidth(c);
            int nw = (int) Math.max(MIN_UNITS, Math.round(w * k));
            sh.setColumnWidth(c, nw);
        }
    }

    /* ===== Общие ячейковые помощники ===== */

    public static Row getOrCreateRow(Sheet sh, int r) {
        Row row = sh.getRow(r);
        return (row != null) ? row : sh.createRow(r);
    }

    public static Cell getOrCreateCell(Row row, int c) {
        Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        return (cell != null) ? cell : row.createCell(c);
    }

    public static CellRangeAddress merge(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress a = new CellRangeAddress(r1, r2, c1, c2);
        sh.addMergedRegion(a);
        return a;
    }

    public static void setCenter(Sheet sh, int r, int c, String text, CellStyle style) {
        Row row = getOrCreateRow(sh, r);
        Cell cell = getOrCreateCell(row, c);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    public static void setText(Row row, int c, String text, CellStyle style) {
        Cell cell = getOrCreateCell(row, c);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    public static void setThinBorder(CellStyle st) {
        st.setBorderTop(BorderStyle.THIN);
        st.setBorderBottom(BorderStyle.THIN);
        st.setBorderLeft(BorderStyle.THIN);
        st.setBorderRight(BorderStyle.THIN);
    }

    public static void setRowHeightCm(Sheet sh, int rowIndex, double cm) {
        getOrCreateRow(sh, rowIndex).setHeightInPoints(cmToPt(cm));
    }

    public static float cmToPt(double cm) {
        return (float) (cm * 72.0 / 2.54);
    }

    public static File ensureXlsx(File f) {
        String name = f.getName().toLowerCase();
        return name.endsWith(".xlsx") ? f : new File(f.getParentFile(), f.getName() + ".xlsx");
    }

    public static String safeDateLine(Map<NoiseTestKind, String> map, NoiseTestKind k) {
        String s = (map != null) ? map.get(k) : null;
        return (s != null && !s.isBlank())
                ? s
                : "Дата, время проведения измерений __.__.____ c __:__ до __:__";
    }

    /** Общая простая «шапка»: строка нумерации + строка даты. */
    public static void writeSimpleHeader(Workbook wb, Sheet sh, String dateLine) {
        org.apache.poi.ss.usermodel.Font f8 = wb.createFont();
        f8.setFontName("Arial");
        f8.setFontHeightInPoints((short) 8);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.setAlignment(HorizontalAlignment.CENTER);
        centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
        centerBorder.setWrapText(false);
        centerBorder.setFont(f8);
        setThinBorder(centerBorder);

        setRowHeightCm(sh, 0, 0.53);
        Row r1 = getOrCreateRow(sh, 0);
        for (int c = 0; c <= 18; c++) setText(r1, c, String.valueOf(c + 1), centerBorder);
        CellRangeAddress t20 = merge(sh, 0, 0, 19, 21); setCenter(sh, 0, 19, "20", centerBorder);
        CellRangeAddress w21 = merge(sh, 0, 0, 22, 24); setCenter(sh, 0, 22, "21", centerBorder);
        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, w21, sh);

        setRowHeightCm(sh, 1, 0.53);
        CellRangeAddress a2y2 = merge(sh, 1, 1, 0, 24);
        Row r2 = getOrCreateRow(sh, 1);
        Cell a = getOrCreateCell(r2, 0);
        a.setCellValue(dateLine);
        a.setCellStyle(centerBorder);
        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, a2y2, sh);
    }

    /** Колонка C (1-я строка блока): "<идентификатор>, <комната>[, <секция>]". Без этажа. */
    public static String formatPlace(Building b, Section sec, Floor f, Space s, Room r) {
        StringBuilder sb = new StringBuilder();
        String spaceId  = (s != null && s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        String roomName = (r != null && r.getName() != null)       ? r.getName().trim()       : "";

        if (!spaceId.isEmpty()) {
            sb.append(spaceId);
            if (!roomName.isEmpty()) sb.append(", ");
        }
        sb.append(roomName);

        boolean multiSections = b != null && b.getSections() != null && b.getSections().size() > 1;
        String secName = (sec != null && sec.getName() != null) ? sec.getName().trim() : "";
        if (multiSections && !secName.isEmpty()) {
            sb.append(", ").append(secName);
        }
        return sb.toString();
    }
}
