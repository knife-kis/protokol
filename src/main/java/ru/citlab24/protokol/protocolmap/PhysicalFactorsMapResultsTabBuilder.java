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

    private PhysicalFactorsMapResultsTabBuilder() {
    }

    static void createResultsSheet(Workbook workbook, List<String> measurementDates) {
        Sheet sheet = workbook.createSheet("Микроклимат");
        applySheetDefaults(workbook, sheet);
        addResultsRows(sheet, measurementDates);
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
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 1);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));
    }

    private static void addResultsRows(Sheet sheet, List<String> measurementDates) {
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

        addMicroclimateHeaderTable(sheet, rowIndex);
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

    private static void addMicroclimateHeaderTable(Sheet sheet, int startRowIndex) {
        Workbook workbook = sheet.getWorkbook();
        Font headerFont = workbook.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 8);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(headerFont);
        headerStyle.setWrapText(true);
        headerStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        Row row1 = sheet.createRow(startRowIndex);
        Row row2 = sheet.createRow(startRowIndex + 1);
        Row row3 = sheet.createRow(startRowIndex + 2);

        mergeAndSet(sheet, startRowIndex, startRowIndex + 1, 0, 0, "№ п/п", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex + 1, 1, 1,
                "Рабочее место, место проведения измерений, цех, участок,\nнаименование профессии или \nдолжности",
                headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex + 1, 2, 2, "Высота от пола, м", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex, 3, 5, "Температура воздуха, ºС", headerStyle);
        mergeAndSet(sheet, startRowIndex + 1, startRowIndex + 1, 3, 5,
                "Измеренная (± расширенная неопределенность)", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex + 1, 6, 6,
                "Температура поверхностей, ºС\nПол. Измеренная (± расширенная неопределенность)", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex, 7, 9, "Результирующая температура,  ºС", headerStyle);
        mergeAndSet(sheet, startRowIndex + 1, startRowIndex + 1, 7, 9,
                "Измеренная (± расширенная неопределенность)", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex, 10, 12, "Относительная влажность воздуха, %", headerStyle);
        mergeAndSet(sheet, startRowIndex + 1, startRowIndex + 1, 10, 12,
                "Измеренная (± расширенная неопределенность)", headerStyle);
        mergeAndSet(sheet, startRowIndex, startRowIndex, 13, 15, "Скорость движения воздуха,  м/с", headerStyle);
        mergeAndSet(sheet, startRowIndex + 1, startRowIndex + 1, 13, 15,
                "Измеренная (± расширенная неопределенность)", headerStyle);

        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 0, 0, "1", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 1, 1, "2", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 2, 2, "3", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 3, 5, "4", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 6, 6, "5", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 7, 9, "6", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 10, 12, "7", headerStyle);
        mergeAndSet(sheet, startRowIndex + 2, startRowIndex + 2, 13, 15, "8", headerStyle);

        applyRowStyle(row1, headerStyle, 0, 15);
        applyRowStyle(row2, headerStyle, 0, 15);
        applyRowStyle(row3, headerStyle, 0, 15);
    }

    private static void mergeAndSet(Sheet sheet,
                                    int firstRow,
                                    int lastRow,
                                    int firstCol,
                                    int lastCol,
                                    String value,
                                    CellStyle style) {
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
        Row row = ensureRow(sheet, firstRow);
        Cell cell = row.createCell(firstCol);
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static void applyRowStyle(Row row, CellStyle style, int firstCol, int lastCol) {
        for (int col = firstCol; col <= lastCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellStyle(style);
        }
    }

    private static Row ensureRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
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
