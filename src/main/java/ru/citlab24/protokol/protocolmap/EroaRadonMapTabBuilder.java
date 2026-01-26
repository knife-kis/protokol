package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

final class EroaRadonMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private EroaRadonMapTabBuilder() {
    }

    static int createEroaRadonResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("ЭРОА Радона");
        applySheetDefaults(workbook, sheet);

        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle borderCenterStyle = createBorderStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        CellStyle borderCenterNoWrapStyle = createBorderNoWrapStyle(workbook, borderCenterStyle);

        int rowIndex = 0;
        Row titleRow = sheet.createRow(rowIndex);
        setCellValue(titleRow, 0,
                "7.5. ЭРОА радона, ЭРОА торона, среднегодовое значение ЭРОА изотопов радона:",
                titleStyle);
        mergeRegionWithoutBorders(sheet, rowIndex, rowIndex, 0, 5);

        rowIndex++;
        Row headerRow = sheet.createRow(rowIndex);
        headerRow.setHeightInPoints(pixelsToPoints(36));
        Row subHeaderRow = sheet.createRow(rowIndex + 1);
        subHeaderRow.setHeightInPoints(pixelsToPoints(65));

        applyMergedHeader(sheet, rowIndex, rowIndex + 1, 0, 0, "№ п/п", borderCenterStyle);
        applyMergedHeader(sheet, rowIndex, rowIndex + 1, 1, 1, "Наименование места\nпроведения измерений",
                borderCenterStyle);
        applyMergedHeader(sheet, rowIndex, rowIndex, 2, 5, "Результаты измерений, Бк/м^3", borderCenterStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 1, 2, 3,
                "Измеренное значение ЭРОА радона (ЭРОА торона)", borderCenterStyle);
        setCellValue(subHeaderRow, 4, "Среднегодовое значение ЭРОА изотопов радона", borderCenterStyle);
        setCellValue(subHeaderRow, 5, "Суммарная неопределенность", borderCenterStyle);

        rowIndex += 2;
        Row numberRow = sheet.createRow(rowIndex);
        setCellValue(numberRow, 0, "1", borderCenterNoWrapStyle);
        setCellValue(numberRow, 1, "2", borderCenterNoWrapStyle);
        applyMergedHeader(sheet, rowIndex, rowIndex, 2, 3, "3", borderCenterNoWrapStyle);
        setCellValue(numberRow, 4, "4", borderCenterNoWrapStyle);
        setCellValue(numberRow, 5, "5", borderCenterNoWrapStyle);

        return rowIndex + 1;
    }

    private static void applySheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP);

        int[] widthsPx = buildColumnWidthsPx();
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

    private static void applyMergedHeader(Sheet sheet,
                                          int rowStart,
                                          int rowEnd,
                                          int colStart,
                                          int colEnd,
                                          String value,
                                          CellStyle style) {
        for (int r = rowStart; r <= rowEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                row = sheet.createRow(r);
            }
            for (int c = colStart; c <= colEnd; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    cell = row.createCell(c);
                }
                cell.setCellStyle(style);
            }
        }
        setCellValue(sheet.getRow(rowStart), colStart, value, style);
        mergeRegion(sheet, rowStart, rowEnd, colStart, colEnd);
    }

    private static void mergeRegion(Sheet sheet, int rowStart, int rowEnd, int colStart, int colEnd) {
        CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
    }

    private static void mergeRegionWithoutBorders(Sheet sheet,
                                                  int rowStart,
                                                  int rowEnd,
                                                  int colStart,
                                                  int colEnd) {
        CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
        sheet.addMergedRegion(region);
    }

    private static CellStyle createTitleStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle createBorderStyle(Workbook workbook,
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

    private static CellStyle createBorderNoWrapStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setWrapText(false);
        return style;
    }

    private static void setCellValue(Row row, int colIndex, String value, CellStyle style) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
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

    private static int[] buildColumnWidthsPx() {
        return new int[]{33, 380, 96, 88, 99, 82};
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
