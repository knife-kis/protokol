package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

final class VentilationMapTabBuilder {
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final float EXTRA_ROW_HEIGHT_POINTS = 15f;

    private VentilationMapTabBuilder() {
    }

    static Sheet createSheet(Workbook workbook) {
        Sheet sheet = workbook.createSheet("Вентиляция");
        CellStyle baseStyle = applySheetDefaults(workbook, sheet);
        int rowIndex = addMergedRow(sheet, 0, 9,
                "7.3. Скорость движения воздуха в рабочих проемах систем вентиляции, кратность воздухообмена",
                baseStyle);
        addTableHeader(sheet, rowIndex, baseStyle);
        return sheet;
    }

    private static CellStyle applySheetDefaults(Workbook workbook, Sheet sheet) {
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
        return baseStyle;
    }

    private static int addMergedRow(Sheet sheet, int rowIndex, int lastCol, String text, CellStyle style) {
        Row row = sheet.createRow(rowIndex);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        if (style != null) {
            cell.setCellStyle(style);
        }
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, lastCol));
        row.setHeightInPoints(12f + EXTRA_ROW_HEIGHT_POINTS);
        return rowIndex + 1;
    }

    private static void addTableHeader(Sheet sheet, int rowIndex, CellStyle baseStyle) {
        Workbook workbook = sheet.getWorkbook();
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.cloneStyleFrom(baseStyle);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.cloneStyleFrom(headerStyle);

        Row headerRow = sheet.createRow(rowIndex);
        headerRow.setHeightInPoints(pixelsToPoints(152));
        createCell(headerRow, 0, "№ п/п", headerStyle);
        createCell(headerRow, 1, "№ точки измерения", headerStyle);
        createCell(headerRow, 2, "Рабочее место, место проведения измерений (Приток/вытяжка)", headerStyle);
        createCell(headerRow, 3, "Измеренные значения скорости воздушного потока (± расширенная неопределенность) м/с",
                headerStyle);
        createCell(headerRow, 4, "", headerStyle);
        createCell(headerRow, 5, "", headerStyle);
        createCell(headerRow, 6, "Площадь сечения\nпроема, м^2", headerStyle);
        createCell(headerRow, 7, "Производительность\nвент системы (канал или общая), м^3/ч", headerStyle);
        createCell(headerRow, 8, "Объем помещения, м^3", headerStyle);
        createCell(headerRow, 9, "Кратность воздухообмена по притоку или вытяжке, ч^-1", headerStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 3, 5));

        Row numberRow = sheet.createRow(rowIndex + 1);
        createCell(numberRow, 0, "1", numberStyle);
        createCell(numberRow, 1, "2", numberStyle);
        createCell(numberRow, 2, "3", numberStyle);
        createCell(numberRow, 3, "4", numberStyle);
        createCell(numberRow, 4, "", numberStyle);
        createCell(numberRow, 5, "", numberStyle);
        createCell(numberRow, 6, "5", numberStyle);
        createCell(numberRow, 7, "6", numberStyle);
        createCell(numberRow, 8, "7", numberStyle);
        createCell(numberRow, 9, "8", numberStyle);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex + 1, rowIndex + 1, 3, 5));
    }

    private static void createCell(Row row, int columnIndex, String text, CellStyle style) {
        Cell cell = row.createCell(columnIndex);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static int[] buildColumnWidthsPx() {
        return new int[]{
                31, 56, 231, 39, 19, 39, 259, 86, 43, 79
        };
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
