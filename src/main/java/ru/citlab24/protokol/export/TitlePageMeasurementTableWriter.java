package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

final class TitlePageMeasurementTableWriter {
    private TitlePageMeasurementTableWriter() {
    }

    static int[][] headerRanges() {
        return new int[][]{
                {0, 3},
                {4, 9},
                {10, 12},
                {13, 19},
                {20, 25}
        };
    }

    static String[] headerTexts() {
        return new String[]{
                "Измеряемый показатель",
                "Наименование, тип средства Измерения",
                "Заводской номер",
                "Погрешность средства измерения",
                "Сведения о поверке (№ свидетельства, срок действия)"
        };
    }

    static void write(Sheet sheet, int headerRow, CellStyle headerStyle, CellStyle dataStyle) {
        String[] headers = headerTexts();

        setMergedText(sheet, headerStyle, headerRow, headerRow, 0, 3, headers[0]);
        setMergedText(sheet, headerStyle, headerRow, headerRow, 4, 9, headers[1]);
        setMergedText(sheet, headerStyle, headerRow, headerRow, 10, 12, headers[2]);
        setMergedText(sheet, headerStyle, headerRow, headerRow, 13, 19, headers[3]);
        setMergedText(sheet, headerStyle, headerRow, headerRow, 20, 25, headers[4]);

        int dataRow = headerRow + 1;
        String indicatorValue = "Длительность интервала времени";
        String instrumentValue = "Секундомеры электронные, Интеграл С-01";
        String serialValue = "462667";
        String errorValue = "Основная абсолютная погрешность, при температуре 25 \u00b1 5 (\u02da\u0421):\n" +
                "\u00b1(9,6\u00b710-6 \u00b7\u0422x+0,01) \u0441\n" +
                "Дополнительная абсолютная погрешность при отклонении температуры от нормальных условий 25 \u00b1 5 (\u02da\u0421) на 1 \u02da\u0421 изменения температуры:\n" +
                "-(2,2\u00b710-6\u00b7\u0422x) \u0441";
        String verificationValue = "\u0421-\u0410\u0428/21-05-2025/433383424 \u0434\u043e 20.05.2026";

        setMergedText(sheet, dataStyle, dataRow, dataRow, 0, 3, indicatorValue);
        setMergedText(sheet, dataStyle, dataRow, dataRow, 4, 9, instrumentValue);
        setMergedText(sheet, dataStyle, dataRow, dataRow, 10, 12, serialValue);
        setMergedText(sheet, dataStyle, dataRow, dataRow, 13, 19, errorValue);
        setMergedText(sheet, dataStyle, dataRow, dataRow, 20, 25, verificationValue);
    }

    private static void setMergedText(Sheet sheet,
                                      CellStyle style,
                                      int firstRow, int lastRow,
                                      int firstCol, int lastCol,
                                      String text) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);

        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(text);
                }
            }
        }
    }
}
