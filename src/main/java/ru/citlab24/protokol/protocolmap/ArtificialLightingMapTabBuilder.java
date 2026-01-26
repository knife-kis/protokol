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

final class ArtificialLightingMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private ArtificialLightingMapTabBuilder() {
    }

    static int createArtificialLightingResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Иск освещение");
        applySheetDefaults(workbook, sheet);

        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle headerVerticalStyle = createHeaderVerticalStyle(workbook, headerStyle);
        CellStyle headerNumberStyle = createHeaderNumberStyle(workbook, headerStyle);

        int rowIndex = 0;
        Row titleRow = sheet.createRow(rowIndex);
        setCellValue(titleRow, 0, "7.11. Результаты измерений световой среды", titleStyle);
        mergeRegionWithoutBorders(sheet, rowIndex, rowIndex, 0, 12);

        rowIndex++;
        Row subTitleRow = sheet.createRow(rowIndex);
        setCellValue(subTitleRow, 0, "7.11.1. Искусственное освещение:", titleStyle);
        mergeRegionWithoutBorders(sheet, rowIndex, rowIndex, 0, 12);

        rowIndex++;
        Row spacerRow = sheet.createRow(rowIndex);
        spacerRow.setHeightInPoints(pixelsToPoints(10));

        Row headerTop = sheet.createRow(rowIndex + 1);
        headerTop.setHeightInPoints(pixelsToPoints(46));
        Row headerMid = sheet.createRow(rowIndex + 2);
        headerMid.setHeightInPoints(pixelsToPoints(59));
        Row headerBottom = sheet.createRow(rowIndex + 3);
        headerBottom.setHeightInPoints(pixelsToPoints(111));

        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 0, 0, "№ п/п", headerStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 1, 1,
                "Наименование места\nпроведения измерений", headerStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 2, 2,
                "Разряд, под разряд зрительной работы", headerVerticalStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 3, 3,
                "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м",
                headerVerticalStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 4, 4, "Вид, тип светильников", headerVerticalStyle);
        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 3, 5, 5, "Число не горящих ламп, шт.", headerVerticalStyle);

        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 1, 6, 9, "Освещенность, лк", headerStyle);
        applyMergedHeader(sheet, rowIndex + 2, rowIndex + 2, 6, 9,
                "Измеренная ± расширенная неопределенность", headerStyle);
        applyMergedHeader(sheet, rowIndex + 3, rowIndex + 3, 6, 8, "Общая", headerStyle);
        setCellValue(headerBottom, 9, "Комбинированная", headerVerticalStyle);
        applyBorderToCell(headerBottom, 9, headerVerticalStyle);

        applyMergedHeader(sheet, rowIndex + 1, rowIndex + 1, 10, 12, "Коэффициент\nпульсации, %", headerStyle);
        applyMergedHeader(sheet, rowIndex + 2, rowIndex + 3, 10, 12,
                "Измеренный\n ± расширенная неопределенность", headerStyle);

        Row numbersRow = sheet.createRow(rowIndex + 4);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 0, 0, "1", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 1, 1, "2", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 2, 2, "3", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 3, 3, "4", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 4, 4, "5", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 5, 5, "6", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 6, 8, "7", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 9, 9, "8", headerNumberStyle);
        applyMergedHeader(sheet, rowIndex + 4, rowIndex + 4, 10, 12, "9", headerNumberStyle);

        numbersRow.setHeightInPoints(pixelsToPoints(17));
        return rowIndex + 5;
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
        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            for (int colIndex = colStart; colIndex <= colEnd; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
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

    private static void applyBorderToCell(Row row, int colIndex, CellStyle style) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellStyle(style);
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

    private static CellStyle createHeaderStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(style);
        return style;
    }

    private static CellStyle createHeaderVerticalStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setRotation((short) 90);
        return style;
    }

    private static CellStyle createHeaderNumberStyle(Workbook workbook, CellStyle baseStyle) {
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
        return new int[]{61, 246, 31, 79, 49, 41, 83, 15, 50, 61, 40, 27, 40};
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
