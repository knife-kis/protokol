package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

final class AreaRadiationMedMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final int LAST_COL = 30;
    private static final int DEFAULT_COLUMN_WIDTH_PX = 33;

    private AreaRadiationMedMapTabBuilder() {
    }

    static Sheet createSheet(Workbook workbook, int profileCount, int controlPointCount) {
        Sheet sheet = workbook.createSheet("МЭД");
        applySheetDefaults(workbook, sheet);

        CellStyle titleStyle = createTextStyle(workbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT, false);
        CellStyle headerCenterStyle = createTextStyle(workbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, true);
        CellStyle dataCenterStyle = createTextStyle(workbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, true);
        CellStyle dataLeftStyle = createTextStyle(workbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT, true);

        int rowIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        setCellValue(row, 0, "7. Результаты измерений на объекте:", titleStyle);
        mergeRegion(sheet, 0, 0, 0, LAST_COL, false);

        row = sheet.createRow(rowIndex++);
        setCellValue(row, 0, "7.9. Результаты измерений МД гамма-излучения", titleStyle);
        mergeRegion(sheet, 1, 1, 0, LAST_COL, false);

        sheet.createRow(rowIndex++);

        float defaultHeight = sheet.getDefaultRowHeightInPoints();
        Row headerRow = sheet.createRow(rowIndex++);
        headerRow.setHeightInPoints(defaultHeight * 2f);
        setCellValue(headerRow, 0, "№ п/п", headerCenterStyle);
        mergeCellRange(sheet, headerRow.getRowNum(), 1, 7,
                "Наименование места проведения измерений", headerCenterStyle);
        mergeCellRange(sheet, headerRow.getRowNum(), 8, LAST_COL,
                "Результаты измерений, мкЗв/ч *10^-2", headerCenterStyle);
        applyBorders(sheet, headerRow.getRowNum(), 0, 0);
        applyBorders(sheet, headerRow.getRowNum(), 1, 7);
        applyBorders(sheet, headerRow.getRowNum(), 8, LAST_COL);

        Row numberRow = sheet.createRow(rowIndex++);
        setCellValue(numberRow, 0, "1", headerCenterStyle);
        mergeCellRange(sheet, numberRow.getRowNum(), 1, 7, "2", headerCenterStyle);
        mergeCellRange(sheet, numberRow.getRowNum(), 8, LAST_COL, "3", headerCenterStyle);
        applyBorders(sheet, numberRow.getRowNum(), 0, 0);
        applyBorders(sheet, numberRow.getRowNum(), 1, 7);
        applyBorders(sheet, numberRow.getRowNum(), 8, LAST_COL);

        int sequence = 1;
        for (int i = 1; i <= profileCount; i++) {
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(defaultHeight * 3f);
            setCellValue(dataRow, 0, String.valueOf(sequence++), dataCenterStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 1, 7, "Профиль " + i, dataLeftStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 8, LAST_COL, "", dataCenterStyle);
            applyBorders(sheet, dataRow.getRowNum(), 0, 0);
            applyBorders(sheet, dataRow.getRowNum(), 1, 7);
            applyBorders(sheet, dataRow.getRowNum(), 8, LAST_COL);
        }
        for (int i = 1; i <= controlPointCount; i++) {
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(defaultHeight * 3f);
            setCellValue(dataRow, 0, String.valueOf(sequence++), dataCenterStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 1, 7, "Контрольная точка " + i, dataLeftStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 8, LAST_COL, "", dataCenterStyle);
            applyBorders(sheet, dataRow.getRowNum(), 0, 0);
            applyBorders(sheet, dataRow.getRowNum(), 1, 7);
            applyBorders(sheet, dataRow.getRowNum(), 8, LAST_COL);
        }

        return sheet;
    }

    private static void applySheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        for (int col = 0; col <= LAST_COL; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(DEFAULT_COLUMN_WIDTH_PX));
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

    private static CellStyle createTextStyle(Workbook workbook,
                                             org.apache.poi.ss.usermodel.HorizontalAlignment alignment,
                                             boolean withBorders) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        if (withBorders) {
            setThinBorders(style);
        }
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

    private static void mergeCellRange(Sheet sheet, int rowIndex, int startCol, int endCol, String value, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(startCol);
        if (cell == null) {
            cell = row.createCell(startCol);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, startCol, endCol);
        sheet.addMergedRegion(region);
        if (style.getBorderBottom() != org.apache.poi.ss.usermodel.BorderStyle.NONE) {
            applyRegionBorders(sheet, region);
        }
    }

    private static void mergeRegion(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, boolean borders) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);
        if (borders) {
            applyRegionBorders(sheet, region);
        }
    }

    private static void applyBorders(Sheet sheet, int rowIndex, int startCol, int endCol) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, startCol, endCol);
        applyRegionBorders(sheet, region);
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
