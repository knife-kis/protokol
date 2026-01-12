package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import ru.citlab24.protokol.tabs.models.*;
import ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind;

import java.io.File;
import java.util.Map;

public class NoiseSheetCommon {

    public static final String NORM_SANPIN_DAY   = "Нормативные требования: СанПиН 1.2.3685-21 (с 07. 00 ч. до 23.00 ч.)";
    public static final String NORM_SANPIN_NIGHT = "Нормативные требования: СанПиН 1.2.3685-21 (с 23. 00 ч. до 07.00 ч.)";
    public static final String NORM_SANPIN       = "Нормативные требования: СанПиН 1.2.3685-21";
    public static final String NORM_SP_51        = "Нормативные требования: СП 51.13330.2011";

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

    /** Заполнение третьей строки блока (T..Y):
     *  - в T и W ставим формулу «1-я строка минус 2-я» с проверкой пустых значений;
     *  - в U и X пишем символ «±» без левых/правых границ;
     *  - в V и Y фиксированное значение «2,3».
     */
    public static void fillNoiseThirdRowDiffs(Sheet sh,
                                              int rowFirst,
                                              int rowSecond,
                                              int rowThird,
                                              CellStyle diffValueStyle,
                                              CellStyle fixedValueStyle,
                                              CellStyle plusMinusStyle) {
        fillDiffFormula(sh, rowFirst, rowSecond, rowThird, 19, diffValueStyle);
        fillDiffFormula(sh, rowFirst, rowSecond, rowThird, 22, diffValueStyle);

        setPlusMinus(sh, rowThird, 20, plusMinusStyle);
        setPlusMinus(sh, rowThird, 23, plusMinusStyle);

        setFixedText(sh, rowThird, 21, "2,3", fixedValueStyle);
        setFixedText(sh, rowThird, 24, "2,3", fixedValueStyle);
    }

    private static void fillDiffFormula(Sheet sh,
                                        int rowFirst,
                                        int rowSecond,
                                        int rowThird,
                                        int column,
                                        CellStyle style) {
        Row r3 = getOrCreateRow(sh, rowThird);
        Cell cell = getOrCreateCell(r3, column);
        cell.setCellStyle(style);

        String col = CellReference.convertNumToColString(column);
        int excelRowFirst = rowFirst + 1;
        int excelRowSecond = rowSecond + 1;
        String refFirst = col + excelRowFirst;
        String refSecond = col + excelRowSecond;
        String diffExpr = refFirst + "-VALUE(" + refSecond + ")";
        String formula = "IF(OR(" + refFirst + "=\"\"," + refSecond + "=\"\"),\"\",ROUND(" +
                diffExpr + ",1))";
        cell.setCellFormula(formula);
    }

    private static void setPlusMinus(Sheet sh, int row, int column, CellStyle style) {
        Row r = getOrCreateRow(sh, row);
        Cell cell = getOrCreateCell(r, column);
        cell.setCellValue("±");
        cell.setCellStyle(style);
    }

    private static void setFixedText(Sheet sh, int row, int column, String text, CellStyle style) {
        Row r = getOrCreateRow(sh, row);
        Cell cell = getOrCreateCell(r, column);
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

    /**
     * Колонка C (1-я строка блока): "<идентификатор>, <комната>[, <секция>]". Без этажа.
     * Для офисов выводим только название комнаты (без идентификатора помещения).
     */
    public static String formatPlace(Building b, Section sec, Floor f, Space s, Room r) {
        StringBuilder sb = new StringBuilder();
        String spaceId  = (s != null && s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        String roomName = (r != null && r.getName() != null)       ? r.getName().trim()       : "";

        boolean officeSpace = s != null && s.getType() == Space.SpaceType.OFFICE;
        boolean appendSpaceId = !officeSpace && !spaceId.isEmpty();

        if (appendSpaceId) {
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
    /** Оценка числа строк при переносе: суммируем по сегментам, разделённым \n. */
    public static int estimateWrappedLines(String text, double colChars) {
        if (text == null || text.isBlank()) return 1;
        colChars = Math.max(1.0, colChars);
        int lines = 0;
        String[] segs = text.split("\\r?\\n");
        for (String seg : segs) {
            int len = Math.max(1, seg.trim().length());
            lines += (int) Math.ceil(len / colChars);
        }
        return Math.max(1, lines);
    }

    /** Подгоняет высоту строки по содержимому указанных столбцов (wrapText=true ОБЯЗАТЕЛЕН в стилях ячеек). */
    public static void adjustRowHeightForWrapped(Sheet sh, int rowIndex, double lineCm, int[] cols, String[] texts) {
        if (cols == null || texts == null || cols.length == 0) return;
        int maxLines = 1;
        for (int i = 0; i < cols.length && i < texts.length; i++) {
            int col = cols[i];
            double colChars = Math.max(1.0, sh.getColumnWidth(col) / 256.0); // «символов» в колонке
            int lines = estimateWrappedLines(texts[i], colChars);
            if (lines > maxLines) maxLines = lines;
        }
        setRowHeightCm(sh, rowIndex, maxLines * lineCm); // 0.53 см на строку, как по ТЗ
    }

    /** Добавляет строку с нормативными требованиями и возвращает следующий индекс строки. */
    public static int appendNormativeRow(Workbook wb,
                                         Sheet sh,
                                         int rowIndex,
                                         String text,
                                         String eqValue,
                                         String maxValue) {
        if (text == null) text = "";
        if (eqValue == null) eqValue = "";
        if (maxValue == null) maxValue = "";

        Font bold = wb.createFont();
        bold.setFontName("Arial");
        bold.setFontHeightInPoints((short) 10);
        bold.setBold(true);

        Font regular = wb.createFont();
        regular.setFontName("Arial");
        regular.setFontHeightInPoints((short) 10);

        CellStyle textStyle = wb.createCellStyle();
        textStyle.setAlignment(HorizontalAlignment.LEFT);
        textStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        textStyle.setWrapText(true);
        textStyle.setFont(bold);
        setThinBorder(textStyle);

        CellStyle centerStyle = wb.createCellStyle();
        centerStyle.setAlignment(HorizontalAlignment.CENTER);
        centerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        centerStyle.setWrapText(false);
        centerStyle.setFont(bold);
        setThinBorder(centerStyle);

        Row row = getOrCreateRow(sh, rowIndex);
        CellRangeAddress normText = merge(sh, rowIndex, rowIndex, 0, 18);
        setText(row, 0, text, textStyle);

        CellRangeAddress eqRange = merge(sh, rowIndex, rowIndex, 19, 21);
        setCenter(sh, rowIndex, 19, eqValue, centerStyle);

        CellRangeAddress maxRange = merge(sh, rowIndex, rowIndex, 22, 24);
        setCenter(sh, rowIndex, 22, maxValue, centerStyle);

        for (CellRangeAddress rg : new CellRangeAddress[]{normText, eqRange, maxRange}) {
            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rg, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rg, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rg, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rg, sh);
        }

        double totalChars = 0.0;
        for (int c = 0; c <= 18; c++) {
            totalChars += Math.max(1.0, sh.getColumnWidth(c) / 256.0);
        }
        int lines = estimateWrappedLines(text, totalChars);
        setRowHeightCm(sh, rowIndex, Math.max(0.53, lines * 0.53));
        return rowIndex + 1;
    }
}
