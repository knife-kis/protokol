package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

final class PhysicalFactorsMapResultsTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final float EXTRA_ROW_HEIGHT_POINTS = 15f;
    private static final float HEADER_ROW_TOP_HEIGHT_POINTS = pixelsToPoints(48);
    private static final float HEADER_ROW_BOTTOM_HEIGHT_POINTS = pixelsToPoints(120);

    private PhysicalFactorsMapResultsTabBuilder() {
    }

    static int createResultsSheet(Workbook workbook, List<String> measurementDates, boolean hasMicroclimateSheet) {
        Sheet sheet = workbook.createSheet("Микроклимат");
        applySheetDefaults(workbook, sheet);
        CellStyle microclimateHeaderStyle = createMicroclimateHeaderStyle(workbook);
        CellStyle microclimateHeaderVerticalStyle = createMicroclimateHeaderVerticalStyle(workbook, microclimateHeaderStyle);
        CellStyle microclimateHeaderNumberStyle = createMicroclimateHeaderNumberStyle(workbook, microclimateHeaderStyle);
        return addResultsRows(sheet, measurementDates, hasMicroclimateSheet,
                microclimateHeaderStyle, microclimateHeaderVerticalStyle, microclimateHeaderNumberStyle);
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

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 0);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(false);
        sheet.setAutobreaks(false);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));
    }

    private static int addResultsRows(Sheet sheet,
                                      List<String> measurementDates,
                                      boolean hasMicroclimateSheet,
                                      CellStyle microclimateHeaderStyle,
                                      CellStyle microclimateHeaderVerticalStyle,
                                      CellStyle microclimateHeaderNumberStyle) {
        List<String> dates = measurementDates == null || measurementDates.isEmpty()
                ? List.of("")
                : measurementDates;

        int rowIndex = 0;
        rowIndex = addMergedRow(sheet, rowIndex, "7. Результаты измерений на объекте: ");

        int sectionIndex = 1;
        for (String date : dates) {
            String textBlock = buildDateBlock(date, sectionIndex);
            rowIndex = addMergedRowWithHeight(sheet, rowIndex, textBlock);
            sectionIndex++;
        }

        if (hasMicroclimateSheet) {
            addMicroclimateHeaderTable(sheet, rowIndex, microclimateHeaderStyle,
                    microclimateHeaderVerticalStyle, microclimateHeaderNumberStyle);
            return rowIndex + 3;
        }
        return rowIndex;
    }

    private static String buildDateBlock(String date, int sectionIndex) {
        String normalizedDate = date == null ? "" : date.trim();
        StringBuilder builder = new StringBuilder();
        builder.append("7.1.").append(sectionIndex).append(" Метеорологические факторы атмосферного воздуха");
        if (!normalizedDate.isEmpty()) {
            builder.append(" ").append(normalizedDate);
        }
        builder.append("\n");
        builder.append("Температура помещения,˚С ").append(longUnderline()).append("\n");
        builder.append("Температура улица,˚С").append(longUnderline()).append("\n");
        builder.append("относительная влажность  помещения, %").append(shortUnderline()).append("\n");
        builder.append("относительная влажность  улица, %").append(longUnderline()).append("\n");
        builder.append("давление помещение, мм рт. ст. ").append(longUnderline()).append("\n");
        builder.append("давление улица, мм рт. ст.").append(longUnderline());
        return builder.toString();
    }

    private static String longUnderline() {
        return "____________________________________________________________________________";
    }

    private static String shortUnderline() {
        String underline = longUnderline();
        return underline.substring(0, underline.length() - 6);
    }

    private static int addMergedRow(Sheet sheet, int rowIndex, String text) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 15));
        row.setHeightInPoints(12f + EXTRA_ROW_HEIGHT_POINTS);
        return rowIndex + 1;
    }

    private static int addMergedRowWithHeight(Sheet sheet, int rowIndex, String text) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 15));

        int lines = Math.max(1, text.split("\\r?\\n").length);
        row.setHeightInPoints(12f * lines + EXTRA_ROW_HEIGHT_POINTS);
        return rowIndex + 1;
    }

    private static void addMicroclimateHeaderTable(Sheet sheet,
                                                   int rowIndex,
                                                   CellStyle headerStyle,
                                                   CellStyle headerVerticalStyle,
                                                   CellStyle headerNumberStyle) {
        Row rowTop = sheet.createRow(rowIndex);
        Row rowBottom = sheet.createRow(rowIndex + 1);
        Row rowNumbers = sheet.createRow(rowIndex + 2);

        rowTop.setHeightInPoints(HEADER_ROW_TOP_HEIGHT_POINTS);
        rowBottom.setHeightInPoints(HEADER_ROW_BOTTOM_HEIGHT_POINTS);

        setMergedCellValue(sheet, rowIndex, rowIndex + 1, 0, 0, "№ п/п", headerStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex + 1, 1, 1,
                "Рабочее место, место проведения измерений, цех, участок,\n" +
                        "наименование профессии или \n" +
                        "должности", headerStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex + 1, 2, 2, "Высота от пола, м", headerVerticalStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex, 3, 5, "Температура воздуха, ºС", headerStyle);
        setMergedCellValue(sheet, rowIndex + 1, rowIndex + 1, 3, 5,
                "Измеренная (± расширенная неопределенность)", headerVerticalStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex + 1, 6, 6,
                "Температура поверхностей, ºС\nПол. Измеренная (± расширенная неопределенность)", headerVerticalStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex, 7, 9, "Результирующая температура,  ºС", headerStyle);
        setMergedCellValue(sheet, rowIndex + 1, rowIndex + 1, 7, 9,
                "Измеренная (± расширенная неопределенность)", headerVerticalStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex, 10, 12, "Относительная влажность воздуха, %", headerStyle);
        setMergedCellValue(sheet, rowIndex + 1, rowIndex + 1, 10, 12,
                "Измеренная (± расширенная неопределенность)", headerVerticalStyle);
        setMergedCellValue(sheet, rowIndex, rowIndex, 13, 15, "Скорость движения воздуха,  м/с", headerStyle);
        setMergedCellValue(sheet, rowIndex + 1, rowIndex + 1, 13, 15,
                "Измеренная (± расширенная неопределенность)", headerVerticalStyle);

        setMergedCellValue(sheet, rowIndex + 2, rowIndex + 2, 3, 5, "4", headerNumberStyle);
        setMergedCellValue(sheet, rowIndex + 2, rowIndex + 2, 7, 9, "6", headerNumberStyle);
        setMergedCellValue(sheet, rowIndex + 2, rowIndex + 2, 10, 12, "7", headerNumberStyle);
        setMergedCellValue(sheet, rowIndex + 2, rowIndex + 2, 13, 15, "8", headerNumberStyle);

        setCellValue(rowNumbers, 0, "1", headerNumberStyle);
        setCellValue(rowNumbers, 1, "2", headerNumberStyle);
        setCellValue(rowNumbers, 2, "3", headerNumberStyle);
        setCellValue(rowNumbers, 6, "5", headerNumberStyle);

        for (int col = 0; col <= 15; col++) {
            Cell cellTop = rowTop.getCell(col);
            if (cellTop == null) {
                cellTop = rowTop.createCell(col);
            }
            if (cellTop.getCellStyle() == null || cellTop.getCellStyle().getIndex() == 0) {
                cellTop.setCellStyle(headerStyle);
            }
            Cell cellBottom = rowBottom.getCell(col);
            if (cellBottom == null) {
                cellBottom = rowBottom.createCell(col);
            }
            if (cellBottom.getCellStyle() == null || cellBottom.getCellStyle().getIndex() == 0) {
                cellBottom.setCellStyle(headerStyle);
            }
            Cell cellNumber = rowNumbers.getCell(col);
            if (cellNumber == null) {
                cellNumber = rowNumbers.createCell(col);
            }
            cellNumber.setCellStyle(headerNumberStyle);
        }
    }

    private static void setMergedCellValue(Sheet sheet,
                                           int rowStart,
                                           int rowEnd,
                                           int colStart,
                                           int colEnd,
                                           String value,
                                           CellStyle style) {
        Row row = sheet.getRow(rowStart);
        if (row == null) {
            row = sheet.createRow(rowStart);
        }
        Cell cell = row.getCell(colStart);
        if (cell == null) {
            cell = row.createCell(colStart);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
        CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
        sheet.addMergedRegion(region);
        applyRegionBorders(sheet, region);
    }

    private static CellStyle createMicroclimateHeaderStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 8);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(style);
        return style;
    }

    private static CellStyle createMicroclimateHeaderVerticalStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setRotation((short) 90);
        return style;
    }

    private static CellStyle createMicroclimateHeaderNumberStyle(Workbook workbook, CellStyle baseStyle) {
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

    private static int[] buildColumnWidthsPx() {
        return new int[]{
                31, 247, 45, 35, 19, 41, 51, 39,
                19, 39, 35, 29, 32, 32, 32, 32
        };
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
