package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

final class MedMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final float MED_HEADER_ROW_HEIGHT_POINTS = pixelsToPoints(58);
    private static final int AREA_MED_LAST_COL = 30;

    private MedMapTabBuilder() {
    }

    static int createMedResultsSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("МЭД");
        applyMedSheetDefaults(workbook, sheet);

        CellStyle titleStyle = createMedTitleStyle(workbook);
        CellStyle borderCenterStyle = createMedBorderStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        CellStyle borderCenterNoWrapStyle = createMedBorderNoWrapStyle(workbook, borderCenterStyle);

        int rowIndex = 0;
        rowIndex = addMedMergedRowWithoutBorders(sheet, rowIndex,
                "7.4. Результаты измерений мощности дозы гамма-излучений ",
                titleStyle);
        rowIndex = addMedMergedRowWithBottomBorder(sheet, rowIndex,
                "Мощность дозы гамма-излучения на открытой местности в пяти точках составила: " +
                        "______________________________________ (мкЗв/ч)",
                titleStyle);

        Row headerRow = sheet.createRow(rowIndex);
        headerRow.setHeightInPoints(MED_HEADER_ROW_HEIGHT_POINTS);
        setMedCellValue(headerRow, 0, "№ п/п", borderCenterStyle);
        setMedCellValue(headerRow, 1, "Наименование места\nпроведения измерений", borderCenterStyle);
        setMedCellValue(headerRow, 2, "Результат измерения (1 этап) мкЗв/ч", borderCenterStyle);
        rowIndex++;

        Row numberRow = sheet.createRow(rowIndex);
        setMedCellValue(numberRow, 0, "1", borderCenterNoWrapStyle);
        setMedCellValue(numberRow, 1, "2", borderCenterNoWrapStyle);
        setMedCellValue(numberRow, 2, "3", borderCenterNoWrapStyle);
        return rowIndex + 1;
    }

    static int createAreaRadiationResultsSheet(Workbook workbook, int profileCount, int controlPointCount) {
        Sheet sheet = workbook.createSheet("МЭД");
        applyAreaRadiationSheetDefaults(workbook, sheet);

        CellStyle titleStyle = createAreaRadiationTextStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT, false);
        CellStyle headerStyle = createAreaRadiationTextStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, true);
        CellStyle headerLeftStyle = createAreaRadiationTextStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT, true);
        CellStyle dataCenterStyle = createAreaRadiationTextStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER, true);
        CellStyle dataLeftStyle = createAreaRadiationTextStyle(workbook, org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT, true);

        int rowIndex = 0;
        Row row = sheet.createRow(rowIndex++);
        setCellValue(row, 0, "7. Результаты измерений на объекте:", titleStyle);
        mergeRegion(sheet, 0, 0, 0, AREA_MED_LAST_COL, false);

        row = sheet.createRow(rowIndex++);
        setCellValue(row, 0, "7.9. Результаты измерений МД гамма-излучения", titleStyle);
        mergeRegion(sheet, 1, 1, 0, AREA_MED_LAST_COL, false);

        sheet.createRow(rowIndex++);

        float defaultHeight = sheet.getDefaultRowHeightInPoints();
        Row headerRow = sheet.createRow(rowIndex++);
        headerRow.setHeightInPoints(defaultHeight * 2f);
        setCellValue(headerRow, 0, "№ п/п", headerStyle);
        mergeCellRange(sheet, headerRow.getRowNum(), 1, 7, "Наименование места проведения измерений", headerLeftStyle);
        mergeCellRange(sheet, headerRow.getRowNum(), 8, AREA_MED_LAST_COL, "Результаты измерений, мкЗв/ч *10^-2", headerStyle);
        applyBorders(sheet, headerRow.getRowNum(), 0, 0);
        applyBorders(sheet, headerRow.getRowNum(), 1, 7);
        applyBorders(sheet, headerRow.getRowNum(), 8, AREA_MED_LAST_COL);

        Row numberRow = sheet.createRow(rowIndex++);
        setCellValue(numberRow, 0, "1", headerStyle);
        mergeCellRange(sheet, numberRow.getRowNum(), 1, 7, "2", headerStyle);
        mergeCellRange(sheet, numberRow.getRowNum(), 8, AREA_MED_LAST_COL, "3", headerStyle);
        applyBorders(sheet, numberRow.getRowNum(), 0, 0);
        applyBorders(sheet, numberRow.getRowNum(), 1, 7);
        applyBorders(sheet, numberRow.getRowNum(), 8, AREA_MED_LAST_COL);

        int sequence = 1;
        for (int i = 1; i <= profileCount; i++) {
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(defaultHeight * 3f);
            setCellValue(dataRow, 0, String.valueOf(sequence++), dataCenterStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 1, 7, "Профиль " + i, dataLeftStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 8, AREA_MED_LAST_COL, "", dataCenterStyle);
            applyBorders(sheet, dataRow.getRowNum(), 0, 0);
            applyBorders(sheet, dataRow.getRowNum(), 1, 7);
            applyBorders(sheet, dataRow.getRowNum(), 8, AREA_MED_LAST_COL);
        }
        for (int i = 1; i <= controlPointCount; i++) {
            Row dataRow = sheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(defaultHeight * 3f);
            setCellValue(dataRow, 0, String.valueOf(sequence++), dataCenterStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 1, 7, "Контрольная точка " + i, dataLeftStyle);
            mergeCellRange(sheet, dataRow.getRowNum(), 8, AREA_MED_LAST_COL, "", dataCenterStyle);
            applyBorders(sheet, dataRow.getRowNum(), 0, 0);
            applyBorders(sheet, dataRow.getRowNum(), 1, 7);
            applyBorders(sheet, dataRow.getRowNum(), 8, AREA_MED_LAST_COL);
        }

        return rowIndex;
    }

    private static void applyMedSheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.TOP);

        int[] widthsPx = buildMedColumnWidthsPx();
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

    private static void applyAreaRadiationSheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        int[] widthsPx = buildAreaRadiationColumnWidthsPx();
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

    private static int addMedMergedRowWithoutBorders(Sheet sheet, int rowIndex, String text, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 2));
        return rowIndex + 1;
    }

    private static int addMedMergedRow(Sheet sheet, int rowIndex, String text, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, 0, 2);
        sheet.addMergedRegion(region);
        applyRegionBorders(sheet, region);
        return rowIndex + 1;
    }

    private static int addMedMergedRowWithBottomBorder(Sheet sheet, int rowIndex, String text, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        for (int col = 0; col <= 2; col++) {
            Cell cell = row.createCell(col);
            if (col == 0) {
                cell.setCellValue(text);
            }
            if (style != null) {
                cell.setCellStyle(style);
            }
        }
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, 0, 2);
        sheet.addMergedRegion(region);
        applyRegionBottomBorder(sheet, region);
        return rowIndex + 1;
    }

    private static CellStyle createMedTitleStyle(Workbook workbook) {
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

    private static CellStyle createMedBorderStyle(Workbook workbook,
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

    private static CellStyle createMedBorderNoWrapStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setWrapText(false);
        return style;
    }

    private static void setMedCellValue(Row row, int colIndex, String value, CellStyle style) {
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

    private static void applyRegionBottomBorder(Sheet sheet, CellRangeAddress region) {
        org.apache.poi.ss.util.RegionUtil.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.NONE, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.NONE, region, sheet);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.NONE, region, sheet);
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

    private static int[] buildMedColumnWidthsPx() {
        return new int[]{33, 168, 750};
    }

    private static int[] buildAreaRadiationColumnWidthsPx() {
        int[] widths = new int[AREA_MED_LAST_COL + 1];
        for (int i = 0; i < widths.length; i++) {
            widths[i] = 33;
        }
        return widths;
    }

    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static CellStyle createAreaRadiationTextStyle(Workbook workbook,
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

    private static double cmToInches(double centimeters) {
        return centimeters / 2.54d;
    }
}
