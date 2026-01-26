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

final class StreetLightingMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private StreetLightingMapTabBuilder() {
    }

    static int createStreetLightingResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Иск освещение (2)");
        applySheetDefaults(workbook, sheet);

        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle headerVerticalStyle = createHeaderVerticalStyle(workbook, headerStyle);
        CellStyle headerNumberStyle = createHeaderNumberStyle(workbook, headerStyle);

        Row titleRow = sheet.createRow(0);
        setCellValue(titleRow, 0,
                "7.11.3 Искусственное освещение. Средняя горизонтальная освещенность на уровне земли",
                titleStyle);
        mergeRegionWithoutBorders(sheet, 0, 0, 0, 5);

        Row spacerRow = sheet.createRow(1);
        spacerRow.setHeightInPoints(pixelsToPoints(10));

        Row headerTop = sheet.createRow(2);
        headerTop.setHeightInPoints(pixelsToPoints(45));
        Row headerMid = sheet.createRow(3);
        Row headerBottom = sheet.createRow(4);
        headerBottom.setHeightInPoints(pixelsToPoints(58));

        applyMergedHeader(sheet, 2, 4, 0, 0, "№ п/п", headerStyle);
        applyMergedHeader(sheet, 2, 4, 1, 1, "Наименование места\nпроведения измерений", headerStyle);
        applyMergedHeader(sheet, 2, 4, 2, 2,
                "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - "
                        + "высота от пола (земли), м",
                headerVerticalStyle);
        applyMergedHeader(sheet, 2, 4, 3, 3, "Вид, тип светильников", headerVerticalStyle);
        applyMergedHeader(sheet, 2, 4, 4, 4, "Число не горящих ламп, шт.", headerVerticalStyle);

        setCellValue(headerTop, 5, "Освещенность", headerStyle);
        applyBorderToCell(headerTop, 5, headerStyle);
        setCellValue(headerMid, 5, "Измеренная", headerStyle);
        applyBorderToCell(headerMid, 5, headerStyle);
        setCellValue(headerBottom, 5,
                "Средняя горизонтальная освещенность на уровне земли, лк",
                headerStyle);
        applyBorderToCell(headerBottom, 5, headerStyle);

        Row numbersRow = sheet.createRow(5);
        numbersRow.setHeightInPoints(pixelsToPoints(112));
        applyMergedHeader(sheet, 5, 5, 0, 0, "1", headerNumberStyle);
        applyMergedHeader(sheet, 5, 5, 1, 1, "2", headerNumberStyle);
        applyMergedHeader(sheet, 5, 5, 2, 2, "3", headerNumberStyle);
        applyMergedHeader(sheet, 5, 5, 3, 3, "4", headerNumberStyle);
        applyMergedHeader(sheet, 5, 5, 4, 4, "5", headerNumberStyle);
        applyMergedHeader(sheet, 5, 5, 5, 5, "6", headerNumberStyle);

        return 6;
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
        return new int[]{62, 245, 79, 48, 41, 485};
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
