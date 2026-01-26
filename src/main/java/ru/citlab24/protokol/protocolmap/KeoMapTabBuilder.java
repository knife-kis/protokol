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

final class KeoMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private KeoMapTabBuilder() {
    }

    static int createKeoResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("КЕО");
        applySheetDefaults(workbook, sheet);

        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle headerVerticalStyle = createHeaderVerticalStyle(workbook, headerStyle);
        CellStyle headerNumberStyle = createHeaderNumberStyle(workbook, headerStyle);

        Row titleRow = sheet.createRow(0);
        setCellValue(titleRow, 0, "7.11.2 Естественное освещение", titleStyle);
        mergeRegionWithoutBorders(sheet, 0, 0, 0, 17);

        Row spacerRow = sheet.createRow(1);
        spacerRow.setHeightInPoints(pixelsToPoints(10));

        Row headerTop = sheet.createRow(2);
        headerTop.setHeightInPoints(pixelsToPoints(49));
        Row headerBottom = sheet.createRow(3);
        headerBottom.setHeightInPoints(pixelsToPoints(178));

        applyMergedHeader(sheet, 2, 3, 0, 0, "№ п/п", headerStyle);
        applyMergedHeader(sheet, 2, 3, 1, 1, "Наименование места\nпроведения измерений", headerStyle);
        applyMergedHeader(sheet, 2, 3, 2, 2,
                "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - "
                        + "высота от пола (земли), м",
                headerVerticalStyle);

        applyMergedHeader(sheet, 2, 2, 3, 5, "При верхнем или комбинированном освещении", headerStyle);
        setCellValue(headerBottom, 3, "Освещенность внутри помещения, лк", headerVerticalStyle);
        applyBorderToCell(headerBottom, 3, headerVerticalStyle);
        setCellValue(headerBottom, 4, "Наружная освещенность, лк", headerVerticalStyle);
        applyBorderToCell(headerBottom, 4, headerVerticalStyle);
        setCellValue(headerBottom, 5, "КЕО, %", headerVerticalStyle);
        applyBorderToCell(headerBottom, 5, headerVerticalStyle);

        applyMergedHeader(sheet, 2, 2, 6, 14, "При боковом освещении", headerStyle);
        applyMergedHeader(sheet, 3, 3, 6, 8,
                "Освещенность внутри помещения ± расширенная неопределенность, лк",
                headerVerticalStyle);
        applyMergedHeader(sheet, 3, 3, 9, 11,
                "Наружная освещенность ± расширенная неопределенность, лк",
                headerVerticalStyle);
        applyMergedHeader(sheet, 3, 3, 12, 14, "КЕО ± расширенная неопределенность, %", headerVerticalStyle);

        applyMergedHeader(sheet, 2, 3, 15, 17,
                "Неравномерность естественного освещения ± расширенная неопределенность",
                headerVerticalStyle);

        Row numbersRow = sheet.createRow(4);
        numbersRow.setHeightInPoints(pixelsToPoints(20));
        applyMergedHeader(sheet, 4, 4, 0, 0, "1", headerNumberStyle);
        applyMergedHeader(sheet, 4, 4, 1, 1, "2", headerNumberStyle);
        applyMergedHeader(sheet, 4, 4, 2, 2, "3", headerNumberStyle);
        applyMergedHeader(sheet, 4, 4, 3, 5, "4", headerNumberStyle);
        applyMergedHeader(sheet, 4, 4, 6, 14, "5", headerNumberStyle);
        applyMergedHeader(sheet, 4, 4, 15, 17, "6", headerNumberStyle);

        return 5;
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
        if (rowStart == rowEnd && colStart == colEnd) {
            return;
        }
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
        if (rowStart == rowEnd && colStart == colEnd) {
            return;
        }
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
        return new int[]{46, 193, 60, 30, 30, 30, 48, 19, 48, 48, 19, 48, 48, 19, 48, 48, 19, 48};
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
