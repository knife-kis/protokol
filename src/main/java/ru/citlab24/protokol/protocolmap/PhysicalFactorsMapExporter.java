package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public final class PhysicalFactorsMapExporter {
    private static final String REGISTRATION_PREFIX = "Регистрационный номер карты замеров:";
    private static final double COLUMN_WIDTH_SCALE = 0.9;
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private PhysicalFactorsMapExporter() {
    }

    public static File generateMap(File sourceFile) throws IOException {
        String registrationNumber = resolveRegistrationNumber(sourceFile);
        MapHeaderData headerData = resolveHeaderData(sourceFile);
        String measurementPerformer = resolveMeasurementPerformer(sourceFile);
        File targetFile = buildTargetFile(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("карта замеров");
            applySheetDefaults(workbook, sheet);
            applyHeaders(sheet, registrationNumber);
            createTitleRows(workbook, sheet, registrationNumber, headerData, measurementPerformer);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                workbook.write(out);
            }
        }

        return targetFile;
    }

    private static String resolveRegistrationNumber(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findRegistrationNumber(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findRegistrationNumber(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.startsWith(REGISTRATION_PREFIX)) {
                    String tail = text.substring(REGISTRATION_PREFIX.length()).trim();
                    if (!tail.isEmpty()) {
                        return tail;
                    }
                    Cell next = row.getCell(cell.getColumnIndex() + 1);
                    if (next != null) {
                        String nextText = formatter.formatCellValue(next).trim();
                        if (!nextText.isEmpty()) {
                            return nextText;
                        }
                    }
                }
            }
        }
        return "";
    }

    private static File buildTargetFile(File sourceFile) {
        String name = sourceFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = dotIndex > 0 ? name.substring(0, dotIndex) : name;
        return new File(sourceFile.getParentFile(), baseName + "_карта.xlsx");
    }

    private static void applySheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 12);
        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);

        int[] widthsPx = buildColumnWidthsPx();
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
            sheet.setDefaultColumnStyle(col, baseStyle);
        }

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

    private static void applyHeaders(Sheet sheet, String registrationNumber) {
        String font = "&\"Arial\"&12";
        Header header = sheet.getHeader();
        header.setLeft(font + "Испытательная лаборатория\nООО «ЦИТ»");
        header.setCenter(font + "Карта замеров № " + registrationNumber + "\nФ8 РИ ИЛ 2-2023");
        header.setRight(font + "\nКоличество страниц: &[Страница] / &[Страниц] \n ");
    }

    private static void createTitleRows(Workbook workbook,
                                        Sheet sheet,
                                        String registrationNumber,
                                        MapHeaderData headerData,
                                        String measurementPerformer) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        Font sectionFont = workbook.createFont();
        sectionFont.setFontName("Arial");
        sectionFont.setFontHeightInPoints((short) 14);
        sectionFont.setBold(true);

        CellStyle sectionStyle = workbook.createCellStyle();
        sectionStyle.setFont(sectionFont);
        sectionStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        sectionStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        sectionStyle.setWrapText(true);

        Font sectionValueFont = workbook.createFont();
        sectionValueFont.setFontName("Arial");
        sectionValueFont.setFontHeightInPoints((short) 12);
        sectionValueFont.setBold(false);

        CellStyle sectionMixedStyle = workbook.createCellStyle();
        sectionMixedStyle.setFont(sectionValueFont);
        sectionMixedStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        sectionMixedStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        sectionMixedStyle.setWrapText(true);

        setMergedCellValue(sheet, 0, "Испытательная лаборатория Общества с ограниченной ответственностью", titleStyle);

        setMergedCellValue(sheet, 1, "«Центр исследовательских технологий»", titleStyle);

        Row spacerRow = sheet.createRow(2);
        spacerRow.setHeightInPoints(pixelsToPoints(3));

        setMergedCellValue(sheet, 3, "КАРТА ЗАМЕРОВ № " + registrationNumber, titleStyle);

        sheet.createRow(4);

        String customerPrefix = "1. Заказчик: ";
        String customerValue = safe(headerData.customerNameAndContacts);
        setMergedCellValueWithPrefix(sheet, 5, customerPrefix, customerValue, sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 5, 0, 31, customerPrefix + customerValue);

        Row heightRow = sheet.createRow(6);
        heightRow.setHeightInPoints(pixelsToPoints(16));

        String datesPrefix = "2. Дата замеров: ";
        String datesValue = safe(headerData.measurementDates);
        setMergedCellValueWithPrefix(sheet, 7, datesPrefix, datesValue, sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 7, 0, 31, datesPrefix + datesValue);

        Row row8 = sheet.createRow(8);
        row8.setHeightInPoints(pixelsToPoints(16));

        String performerText = "3. Измерения провел, подпись: " + safe(measurementPerformer);
        setMergedCellValue(sheet, 9, performerText, sectionStyle);

        Row row10 = sheet.createRow(10);
        row10.setHeightInPoints(pixelsToPoints(16));

        String representativePrefix = "4. Измерения проведены в присутствии представителя: ";
        String representativeValue = safe(headerData.representative);
        setMergedCellValueWithPrefix(sheet, 11, representativePrefix, representativeValue,
                sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 11, 0, 31, representativePrefix + representativeValue);

        Row row12 = sheet.createRow(12);
        row12.setHeightInPoints(pixelsToPoints(16));

        setMergedCellValue(sheet, 13, "Подпись:_______________________________", sectionStyle);

        Row row14 = sheet.createRow(14);
        row14.setHeightInPoints(pixelsToPoints(16));

        String controlSuffix = resolveControlSuffix(measurementPerformer);
        String controlText = "Контроль ведения записей осуществлен: " + controlSuffix;
        setMergedCellValue(sheet, 15, controlText, sectionStyle);

        Row row16 = sheet.createRow(16);
        row16.setHeightInPoints(pixelsToPoints(16));

        setMergedCellValue(sheet, 17,
                "Результат контроля:____________________________________________",
                sectionStyle);
    }

    private static void setMergedCellValue(Sheet sheet, int rowIndex, String text, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            cell = row.createCell(0);
        }
        cell.setCellStyle(style);
        cell.setCellValue(text);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 31));
    }

    private static void setMergedCellValueWithPrefix(Sheet sheet,
                                                     int rowIndex,
                                                     String prefix,
                                                     String value,
                                                     Font prefixFont,
                                                     Font valueFont,
                                                     CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            cell = row.createCell(0);
        }

        String text = (prefix == null ? "" : prefix) + safe(value);
        org.apache.poi.ss.usermodel.RichTextString richText =
                cell.getSheet().getWorkbook().getCreationHelper().createRichTextString(text);
        int prefixLength = prefix == null ? 0 : prefix.length();
        richText.applyFont(0, prefixLength, prefixFont);
        richText.applyFont(prefixLength, text.length(), valueFont);
        cell.setCellStyle(style);
        cell.setCellValue(richText);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 31));
    }

    private static int[] buildColumnWidthsPx() {
        int[] baseWidths = new int[32];
        baseWidths[0] = 36;
        baseWidths[1] = 33;
        baseWidths[2] = 32;
        baseWidths[3] = 32;
        baseWidths[4] = 32;
        baseWidths[5] = 32;
        baseWidths[6] = 32;
        baseWidths[7] = 32;
        baseWidths[8] = 32;
        for (int col = 9; col <= 17; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[18] = 53;
        baseWidths[19] = 41;
        baseWidths[20] = 58;
        baseWidths[21] = 32;
        for (int col = 22; col <= 29; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[30] = 36;
        baseWidths[31] = 11;

        int[] widths = new int[baseWidths.length];
        for (int i = 0; i < baseWidths.length; i++) {
            widths[i] = (int) Math.round(baseWidths[i] * COLUMN_WIDTH_SCALE);
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

    private static double cmToInches(double centimeters) {
        return centimeters / 2.54d;
    }

    private static float pixelsToPoints(int pixels) {
        return pixels * 0.75f;
    }

    private static float pointsToPixels(float points) {
        return points / 0.75f;
    }

    private static MapHeaderData resolveHeaderData(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new MapHeaderData("", "", "");
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return new MapHeaderData("", "", "");
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findHeaderData(sheet);
        } catch (Exception ex) {
            return new MapHeaderData("", "", "");
        }
    }

    private static MapHeaderData findHeaderData(Sheet sheet) {
        String customer = "";
        String dates = "";
        String representative = "";
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (customer.isEmpty()) {
                    customer = extractCustomer(text);
                    if (customer.isEmpty() && text.startsWith(CUSTOMER_PREFIX)) {
                        customer = readNextCellText(row, cell, formatter);
                    }
                }
                if (dates.isEmpty()) {
                    dates = extractMeasurementDates(text);
                }
                if (representative.isEmpty()) {
                    representative = extractRepresentative(text);
                    if (representative.isEmpty() && text.startsWith(REPRESENTATIVE_PREFIX)) {
                        representative = readNextCellText(row, cell, formatter);
                    }
                }
                if (!customer.isEmpty() && !dates.isEmpty() && !representative.isEmpty()) {
                    return new MapHeaderData(customer, dates, representative);
                }
            }
        }
        return new MapHeaderData(customer, dates, representative);
    }

    private static String extractCustomer(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(CUSTOMER_PREFIX);
        if (index < 0) {
            return "";
        }
        String tail = text.substring(index + CUSTOMER_PREFIX.length()).trim();
        return tail;
    }

    private static String extractMeasurementDates(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        int start = text.indexOf(MEASUREMENT_DATES_PHRASE);
        if (start < 0) {
            return "";
        }
        int from = start + MEASUREMENT_DATES_PHRASE.length();
        String tail = text.substring(from).trim();
        if (tail.isEmpty()) {
            return "";
        }
        java.util.LinkedHashSet<String> dates = new java.util.LinkedHashSet<>();
        java.util.regex.Matcher matcher = DATE_PATTERN.matcher(tail);
        while (matcher.find()) {
            dates.add(matcher.group());
        }
        if (dates.isEmpty()) {
            return "";
        }
        return String.join(", ", dates);
    }

    private static String extractRepresentative(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(REPRESENTATIVE_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + REPRESENTATIVE_PREFIX.length()).trim();
    }

    private static String readNextCellText(Row row, Cell cell, DataFormatter formatter) {
        Cell next = row.getCell(cell.getColumnIndex() + 1);
        if (next == null) {
            return "";
        }
        String nextText = formatter.formatCellValue(next).trim();
        return nextText;
    }

    private static String resolveMeasurementPerformer(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            for (int idx = workbook.getNumberOfSheets() - 1; idx >= 0; idx--) {
                Sheet sheet = workbook.getSheetAt(idx);
                if (sheet == null) {
                    continue;
                }
                if (isGeneratorSheet(sheet.getSheetName())) {
                    continue;
                }
                String performer = findMeasurementPerformer(sheet, formatter);
                if (!performer.isEmpty()) {
                    return performer;
                }
            }
            return "";
        } catch (Exception ex) {
            return "";
        }
    }

    private static boolean isGeneratorSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        return sheetName.equalsIgnoreCase("генератор")
                || sheetName.equalsIgnoreCase("генератор (2.0.)");
    }

    private static String findMeasurementPerformer(Sheet sheet, DataFormatter formatter) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.contains("Измерения проводил:")) {
                    if (text.contains("Тарновский")) {
                        return "Тарновский М.О.";
                    }
                    if (text.contains("Белов")) {
                        return "Белов Д.А.";
                    }
                }
            }
        }
        return "";
    }

    private static String resolveControlSuffix(String measurementPerformer) {
        if ("Тарновский М.О.".equals(measurementPerformer)) {
            return "Инженер Белов Д.А.";
        }
        if ("Белов Д.А.".equals(measurementPerformer)) {
            return "Заведующий лабораторией Тарновский М.О.";
        }
        return "";
    }

    private static String safe(String value) {
        return value == null ? "" : value.trim();
    }

    private static void adjustRowHeightForMergedTextDoubling(Sheet sheet,
                                                             int rowIndex,
                                                             int firstCol,
                                                             int lastCol,
                                                             String text) {
        if (sheet == null) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double totalChars = totalColumnChars(sheet, firstCol, lastCol);
        int lines = estimateWrappedLines(text, totalChars);

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        int multiplier = lines > 1 ? 2 : 1;

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * multiplier)));
    }

    private static double totalColumnChars(Sheet sheet, int firstCol, int lastCol) {
        double totalChars = 0.0;
        for (int c = firstCol; c <= lastCol; c++) {
            totalChars += sheet.getColumnWidth(c) / 256.0;
        }
        return Math.max(1.0, totalChars);
    }

    private static int estimateWrappedLines(String text, double colChars) {
        if (text == null || text.isBlank()) {
            return 1;
        }

        int lines = 0;
        String[] segments = text.split("\\r?\\n");
        for (String seg : segments) {
            int len = Math.max(1, seg.trim().length());
            lines += (int) Math.ceil(len / Math.max(1.0, colChars));
        }
        return Math.max(1, lines);
    }

    private static final String CUSTOMER_PREFIX =
            "Наименование и контактные данные заявителя (заказчика):";
    private static final String MEASUREMENT_DATES_PHRASE = "Измерения были проведены";
    private static final String REPRESENTATIVE_PREFIX =
            "Измерения проводились в присутствии представителя заказчика:";
    private static final java.util.regex.Pattern DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{2}\\.\\d{2}\\.\\d{4}\\b");

    private static final class MapHeaderData {
        private final String customerNameAndContacts;
        private final String measurementDates;
        private final String representative;

        private MapHeaderData(String customerNameAndContacts,
                              String measurementDates,
                              String representative) {
            this.customerNameAndContacts = customerNameAndContacts;
            this.measurementDates = measurementDates;
            this.representative = representative;
        }
    }
}
