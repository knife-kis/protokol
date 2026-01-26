package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

final class Med3MapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private Med3MapTabBuilder() {
    }

    static int createMed3ResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("МЭД (3)");
        applyMed3SheetDefaults(workbook, sheet);

        CellStyle borderCenterStyle = createMed3BorderStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        CellStyle borderCenterNoWrapStyle = createMed3BorderNoWrapStyle(workbook, borderCenterStyle);

        int rowIndex = 0;
        Row headerRow = sheet.createRow(rowIndex);
        headerRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints() * 2f);
        setMed3CellValue(headerRow, 0, "№ п/п", borderCenterStyle);
        setMed3CellValue(headerRow, 1, "Наименование места\nпроведения измерений", borderCenterStyle);
        for (int col = 2; col <= 4; col++) {
            setMed3CellValue(headerRow, col, col == 2
                    ? "Мощность дозы гамма-\nизлучения ± ΔН, мкЗв/ч"
                    : "", borderCenterStyle);
        }
        CellRangeAddress headerRegion = new CellRangeAddress(rowIndex, rowIndex, 2, 4);
        sheet.addMergedRegion(headerRegion);
        applyRegionBorders(sheet, headerRegion);
        rowIndex++;

        Row numberRow = sheet.createRow(rowIndex);
        setMed3CellValue(numberRow, 0, "1", borderCenterNoWrapStyle);
        setMed3CellValue(numberRow, 1, "2", borderCenterNoWrapStyle);
        for (int col = 2; col <= 4; col++) {
            setMed3CellValue(numberRow, col, col == 2 ? "3" : "", borderCenterNoWrapStyle);
        }
        CellRangeAddress numberRegion = new CellRangeAddress(rowIndex, rowIndex, 2, 4);
        sheet.addMergedRegion(numberRegion);
        applyRegionBorders(sheet, numberRegion);
        return rowIndex + 1;
    }

    private static void applyMed3SheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP);

        int[] widthsPx = buildMed3ColumnWidthsPx();
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
            sheet.setDefaultColumnStyle(col, baseStyle);
        }
        sheet.setDefaultRowHeightInPoints(pixelsToPoints(17));

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));
    }

    private static CellStyle createMed3BorderStyle(Workbook workbook,
                                                   org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(style);
        return style;
    }

    private static CellStyle createMed3BorderNoWrapStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setWrapText(false);
        return style;
    }

    private static void setMed3CellValue(Row row, int colIndex, String value, CellStyle style) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void applyRegionBorders(Sheet sheet, CellRangeAddress region) {
        org.apache.poi.ss.util.RegionUtil.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
    }

    private static void setThinBorders(CellStyle style) {
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    }

    private static float pixelsToPoints(float pixels) {
        return pixels * 0.75f;
    }

    private static int[] buildMed3ColumnWidthsPx() {
        return new int[]{33, 550, 74, 20, 74};
    }

    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static double cmToInches(double centimeters) {
        return centimeters / 2.54d;
    }
}
