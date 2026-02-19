package ru.citlab24.protokol.tabs.qms;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

final class ShewhartMapExcelExporter {

    private static final String SHEET_NAME = "Карта Шухарта";
    private static final int HEADER_ROW_1 = 0;
    private static final int HEADER_ROW_2 = 1;
    private static final int COL_A = 0;
    private static final int COL_B = 1;
    private static final int COL_AG = 32;

    private static final int[] FREQUENCIES = {
            100, 125, 160, 200, 250, 315, 400, 500,
            630, 800, 1000, 1250, 1600, 2000, 2500, 3150
    };

    private ShewhartMapExcelExporter() {
    }

    static void exportStaticTitle(File targetFile, List<File> inputFiles) throws IOException {
        if (inputFiles == null || inputFiles.isEmpty()) {
            throw new IllegalArgumentException("Для преобразования необходим хотя бы один входной файл.");
        }

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet(SHEET_NAME);

            CellStyle headerStyle = createHeaderStyle(workbook);

            Row row1 = getOrCreateRow(sheet, HEADER_ROW_1);
            Row row2 = getOrCreateRow(sheet, HEADER_ROW_2);

            for (int col = COL_A; col <= COL_AG; col++) {
                Cell cell1 = row1.createCell(col);
                Cell cell2 = row2.createCell(col);
                cell1.setCellStyle(headerStyle);
                cell2.setCellStyle(headerStyle);
            }

            CellRangeAddress aHeaderRegion = new CellRangeAddress(HEADER_ROW_1, HEADER_ROW_2, COL_A, COL_A);
            sheet.addMergedRegion(aHeaderRegion);

            CellRangeAddress titleRegion = new CellRangeAddress(HEADER_ROW_1, HEADER_ROW_1, COL_B, COL_AG);
            sheet.addMergedRegion(titleRegion);
            row1.getCell(COL_B).setCellValue("Третьоктавные полосы");

            for (int i = 0; i < FREQUENCIES.length; i++) {
                int firstCol = COL_B + i * 2;
                int secondCol = firstCol + 1;
                CellRangeAddress pairRegion = new CellRangeAddress(HEADER_ROW_2, HEADER_ROW_2, firstCol, secondCol);
                sheet.addMergedRegion(pairRegion);
                row2.getCell(firstCol).setCellValue(FREQUENCIES[i]);
            }

            for (int col = COL_A; col <= COL_AG; col++) {
                sheet.setColumnWidth(col, 12 * 256);
            }
            sheet.setColumnWidth(COL_A, 14 * 256);
            row1.setHeightInPoints(22f);
            row2.setHeightInPoints(22f);

            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                workbook.write(outputStream);
            }
        }
    }

    private static Row getOrCreateRow(org.apache.poi.ss.usermodel.Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? sheet.createRow(rowIndex) : row;
    }

    private static CellStyle createHeaderStyle(XSSFWorkbook workbook) {
        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }
}
