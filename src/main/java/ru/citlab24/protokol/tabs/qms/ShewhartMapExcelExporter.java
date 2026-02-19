package ru.citlab24.protokol.tabs.qms;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class ShewhartMapExcelExporter {

    private static final String SHEET_NAME = "Карта Шухарта";
    private static final int HEADER_ROW_1 = 0;
    private static final int HEADER_ROW_2 = 1;
    private static final int FIRST_DATA_ROW = 2;
    private static final int COL_A = 0;
    private static final int COL_B = 1;
    private static final int COL_AG = 32;
    private static final int DATA_COLUMN_WIDTH_PX = 45;
    private static final int SOURCE_FIRST_DATA_COL = 1; // B
    private static final int SOURCE_LAST_DATA_COL = 16; // Q
    private static final Pattern FILENAME_DATE_PATTERN = Pattern.compile("\\b\\d{2}\\.(\\d{2})\\.(?:\\d{2}|\\d{4})\\b");
    private static final String[] MONTHS_RU = {
            "январь", "февраль", "март", "апрель", "май", "июнь",
            "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"
    };

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
            Sheet sheet = workbook.createSheet(SHEET_NAME);

            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle bodyStyle = createBodyStyle(workbook);
            CellStyle numericBodyStyle = createNumericBodyStyle(workbook, bodyStyle);
            DataFormatter formatter = new DataFormatter(Locale.ROOT);

            createHeader(sheet, headerStyle);
            sheet.getPrintSetup().setLandscape(true);

            int targetRowIndex = FIRST_DATA_ROW;
            for (File inputFile : inputFiles) {
                String monthFromFileName = extractMonthFromFileName(inputFile.getName());
                targetRowIndex = appendFileData(sheet, bodyStyle, numericBodyStyle, formatter, inputFile, monthFromFileName, targetRowIndex);
            }

            for (int col = COL_B; col <= COL_AG; col++) {
                sheet.setColumnWidth(col, pixelsToWidthUnits(DATA_COLUMN_WIDTH_PX));
            }
            sheet.setColumnWidth(COL_A, 14 * 256);

            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void createHeader(Sheet sheet, CellStyle headerStyle) {
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

        row1.setHeightInPoints(22f);
        row2.setHeightInPoints(22f);
    }

    private static int appendFileData(
            Sheet targetSheet,
            CellStyle bodyStyle,
            CellStyle numericBodyStyle,
            DataFormatter formatter,
            File inputFile,
            String monthFromFileName,
            int startTargetRow
    ) throws IOException {
        int targetRowIndex = startTargetRow;

        try (FileInputStream stream = new FileInputStream(inputFile);
             Workbook sourceWorkbook = WorkbookFactory.create(stream)) {

            for (int rwIndex = 1; ; rwIndex++) {
                String firstSheetName = "RW" + rwIndex + "-1";
                String secondSheetName = "RW" + rwIndex + "-2";

                Sheet firstSheet = sourceWorkbook.getSheet(firstSheetName);
                Sheet secondSheet = sourceWorkbook.getSheet(secondSheetName);
                if (firstSheet == null && secondSheet == null) {
                    break;
                }

                Row targetRow = getOrCreateRow(targetSheet, targetRowIndex);
                fillRowStyle(targetRow, bodyStyle);
                targetRow.getCell(COL_A).setCellValue(monthFromFileName);

                copyRwValues(firstSheet, formatter, targetRow, true, numericBodyStyle);
                copyRwValues(secondSheet, formatter, targetRow, false, numericBodyStyle);

                targetRowIndex++;
            }
        }

        return targetRowIndex;
    }

    private static void fillRowStyle(Row row, CellStyle style) {
        for (int col = COL_A; col <= COL_AG; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellStyle(style);
        }
    }

    private static void copyRwValues(
            Sheet sourceSheet,
            DataFormatter formatter,
            Row targetRow,
            boolean oddColumns,
            CellStyle numericBodyStyle
    ) {
        if (sourceSheet == null) {
            return;
        }

        Row sourceRow = findRdbRow(sourceSheet, formatter);
        if (sourceRow == null) {
            return;
        }

        for (int sourceCol = SOURCE_FIRST_DATA_COL; sourceCol <= SOURCE_LAST_DATA_COL; sourceCol++) {
            Cell sourceCell = sourceRow.getCell(sourceCol);
            int frequencyIndex = sourceCol - SOURCE_FIRST_DATA_COL;
            int targetCol = COL_B + frequencyIndex * 2 + (oddColumns ? 0 : 1);
            setCellValue(targetRow.getCell(targetCol), sourceCell, formatter, numericBodyStyle);
        }
    }

    private static Row findRdbRow(Sheet sheet, DataFormatter formatter) {
        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            String value = formatter.formatCellValue(row.getCell(COL_A));
            if ("R, дБ".equalsIgnoreCase(normalize(value))) {
                return row;
            }
        }
        return null;
    }

    private static void setCellValue(Cell targetCell, Cell sourceCell, DataFormatter formatter, CellStyle numericBodyStyle) {
        if (sourceCell == null) {
            targetCell.setBlank();
            return;
        }

        if (sourceCell.getCellType() == CellType.NUMERIC) {
            targetCell.setCellValue(sourceCell.getNumericCellValue());
            targetCell.setCellStyle(numericBodyStyle);
            return;
        }

        String value = formatter.formatCellValue(sourceCell).trim();
        if (value.isEmpty()) {
            targetCell.setBlank();
            return;
        }

        try {
            targetCell.setCellValue(Double.parseDouble(value.replace(',', '.')));
            targetCell.setCellStyle(numericBodyStyle);
        } catch (NumberFormatException ex) {
            targetCell.setCellValue(value);
        }
    }

    private static String extractMonthFromFileName(String fileName) {
        Matcher matcher = FILENAME_DATE_PATTERN.matcher(fileName);
        if (matcher.find()) {
            int monthNumber = Integer.parseInt(matcher.group(1));
            if (monthNumber >= 1 && monthNumber <= MONTHS_RU.length) {
                return MONTHS_RU[monthNumber - 1];
            }
        }
        return "";
    }

    private static int pixelsToWidthUnits(int pixels) {
        return Math.max(0, (int) Math.round(((pixels - 5d) / 7d) * 256d));
    }

    private static String normalize(String value) {
        return value == null ? "" : value.replace('\u00A0', ' ').trim();
    }

    private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
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

    private static CellStyle createBodyStyle(XSSFWorkbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        return style;
    }

    private static CellStyle createNumericBodyStyle(XSSFWorkbook workbook, CellStyle baseStyle) {
        CellStyle numericStyle = workbook.createCellStyle();
        numericStyle.cloneStyleFrom(baseStyle);
        numericStyle.setDataFormat(workbook.createDataFormat().getFormat("0.0"));
        return numericStyle;
    }
}
