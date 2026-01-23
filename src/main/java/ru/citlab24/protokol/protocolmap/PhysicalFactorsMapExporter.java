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
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

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
        String protocolNumber = resolveProtocolNumber(sourceFile);
        String contractText = resolveContractText(sourceFile);
        String measurementPerformer = resolveMeasurementPerformer(sourceFile);
        String controlDate = resolveControlDate(sourceFile);
        String specialConditions = resolveSpecialConditions(sourceFile);
        String measurementMethods = resolveMeasurementMethods(sourceFile);
        java.util.List<InstrumentData> instruments = resolveMeasurementInstruments(sourceFile);
        File targetFile = buildTargetFile(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("карта замеров");
            applySheetDefaults(workbook, sheet);
            applyHeaders(sheet, registrationNumber);
            createTitleRows(workbook, sheet, registrationNumber, headerData, measurementPerformer, controlDate);
            createSecondPageRows(workbook, sheet, protocolNumber, contractText, headerData,
                    specialConditions, measurementMethods, instruments);

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
                                        String measurementPerformer,
                                        String controlDate) {
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

        String performerPrefix = "3. Измерения провел, подпись: ";
        String performerValue = safe(measurementPerformer);
        setMergedCellValueWithPrefix(sheet, 9, performerPrefix, performerValue,
                sectionFont, sectionValueFont, sectionMixedStyle);

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

        String controlPrefix = "5. Контроль ведения записей осуществлен: ";
        String controlSuffix = resolveControlSuffix(measurementPerformer);
        setMergedCellValueWithPrefix(sheet, 15, controlPrefix, controlSuffix,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row row16 = sheet.createRow(16);
        row16.setHeightInPoints(pixelsToPoints(16));

        String controlResultPrefix = "Результат контроля: ";
        String controlResultValue = "соответствует/не соответствует";
        setMergedCellValueWithPrefix(sheet, 17, controlResultPrefix, controlResultValue,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row row18 = sheet.createRow(18);
        row18.setHeightInPoints(pixelsToPoints(16));

        String controlDatePrefix = "Дата контроля: ";
        setMergedCellValueWithPrefix(sheet, 19, controlDatePrefix, controlDate,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row spacerAfterControlDate = sheet.createRow(20);
        spacerAfterControlDate.setHeightInPoints(pixelsToPoints(16));
    }

    private static void createSecondPageRows(Workbook workbook,
                                             Sheet sheet,
                                             String protocolNumber,
                                             String contractText,
                                             MapHeaderData headerData,
                                             String specialConditions,
                                             String measurementMethods,
                                             java.util.List<InstrumentData> instruments) {
        int startRow = 21;
        sheet.setRowBreak(startRow - 1);

        Font plainFont = workbook.createFont();
        plainFont.setFontName("Arial");
        plainFont.setFontHeightInPoints((short) 12);
        plainFont.setBold(false);

        CellStyle plainStyle = workbook.createCellStyle();
        plainStyle.setFont(plainFont);
        plainStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        plainStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        plainStyle.setWrapText(true);

        int rowIndex = startRow;
        setMergedCellValue(sheet, rowIndex,
                "1. Номер протокола " + safe(protocolNumber), plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "2. Договор " + safe(contractText), plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "3. Наименование и контактные данные Заказчика: " + safe(headerData.customerNameAndContacts),
                plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "Юридический адрес заказчика: " + safe(headerData.customerLegalAddress), plainStyle);
        rowIndex++;
        String objectNameText = "4. Наименование объекта: " + safe(headerData.objectName);
        setMergedCellValue(sheet, rowIndex, objectNameText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, objectNameText);
        rowIndex++;

        String objectAddressText = "Адрес объекта " + safe(headerData.objectAddress);
        setMergedCellValue(sheet, rowIndex, objectAddressText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, objectAddressText);
        rowIndex++;

        setMergedCellValue(sheet, rowIndex, "5. Дополнительные сведения", plainStyle);
        rowIndex++;

        String specialConditionsText = "5.1. Особые условия: " + safe(specialConditions);
        setMergedCellValue(sheet, rowIndex, specialConditionsText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, specialConditionsText);
        rowIndex++;

        String measurementMethodsText = "5.2. Методы измерения " + safe(measurementMethods);
        setMergedCellValue(sheet, rowIndex, measurementMethodsText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, measurementMethodsText);
        rowIndex++;

        setMergedCellValue(sheet, rowIndex,
                "5.3. Приборы для измерения (используемое отметить):", plainStyle);
        rowIndex++;

        Font tableHeaderFont = workbook.createFont();
        tableHeaderFont.setFontName("Arial");
        tableHeaderFont.setFontHeightInPoints((short) 12);
        tableHeaderFont.setBold(true);

        CellStyle tableHeaderStyle = workbook.createCellStyle();
        tableHeaderStyle.setFont(tableHeaderFont);
        tableHeaderStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        tableHeaderStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        tableHeaderStyle.setWrapText(true);
        setThinBorders(tableHeaderStyle);

        CellStyle tableCellStyle = workbook.createCellStyle();
        tableCellStyle.setFont(plainFont);
        tableCellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        tableCellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        tableCellStyle.setWrapText(true);
        setThinBorders(tableCellStyle);

        CellStyle checkboxStyle = workbook.createCellStyle();
        checkboxStyle.setFont(plainFont);
        checkboxStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        checkboxStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(checkboxStyle);

        rowIndex = addInstrumentRow(sheet, rowIndex, "Наименование", "зав. №", "☑",
                tableHeaderStyle, tableHeaderStyle, checkboxStyle);

        if (instruments != null) {
            for (InstrumentData instrument : instruments) {
                rowIndex = addInstrumentRow(sheet, rowIndex, instrument.name, instrument.serialNumber, "☑",
                        tableCellStyle, tableCellStyle, checkboxStyle);
            }
        }
    }

    private static int addInstrumentRow(Sheet sheet,
                                        int rowIndex,
                                        String name,
                                        String serialNumber,
                                        String checkbox,
                                        CellStyle nameStyle,
                                        CellStyle serialStyle,
                                        CellStyle checkboxStyle) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        mergeCellRangeWithStyle(sheet, rowIndex, 1, 19, nameStyle, safe(name));
        mergeCellRangeWithStyle(sheet, rowIndex, 20, 29, serialStyle, safe(serialNumber));

        Cell checkboxCell = row.getCell(30);
        if (checkboxCell == null) {
            checkboxCell = row.createCell(30);
        }
        checkboxCell.setCellStyle(checkboxStyle);
        checkboxCell.setCellValue(safe(checkbox));

        adjustRowHeightForInstrumentRow(sheet, rowIndex, safe(name), safe(serialNumber));

        rowIndex++;
        return rowIndex;
    }

    private static void mergeCellRangeWithStyle(Sheet sheet,
                                                int rowIndex,
                                                int firstCol,
                                                int lastCol,
                                                CellStyle style,
                                                String value) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol);
        sheet.addMergedRegion(region);

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        // Проставляем стиль во всех ячейках диапазона, а значение — в первой
        for (int col = firstCol; col <= lastCol; col++) {
            Cell c = row.getCell(col);
            if (c == null) c = row.createCell(col);
            c.setCellStyle(style);
            if (col == firstCol) {
                c.setCellValue(value);
            }
        }

        // Для merged-регионов границы надёжнее “прожимать” через RegionUtil
        RegionUtil.setBorderTop(style.getBorderTop(), region, sheet);
        RegionUtil.setBorderBottom(style.getBorderBottom(), region, sheet);
        RegionUtil.setBorderLeft(style.getBorderLeft(), region, sheet);
        RegionUtil.setBorderRight(style.getBorderRight(), region, sheet);
    }

    private static void setThinBorders(CellStyle style) {
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
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
            return new MapHeaderData("", "", "", "", "", "");
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return new MapHeaderData("", "", "", "", "", "");
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findHeaderData(sheet);
        } catch (Exception ex) {
            return new MapHeaderData("", "", "", "", "", "");
        }
    }


    private static MapHeaderData findHeaderData(Sheet sheet) {
        String customer = "";
        String dates = "";
        String representative = "";
        String legalAddress = "";
        String objectName = "";
        String objectAddress = "";
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell);
                String text = rawText.trim();
                String normalized = normalizeText(rawText);
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
                if (legalAddress.isEmpty()) {
                    legalAddress = extractLegalAddress(normalized);
                    if (legalAddress.isEmpty() && normalized.startsWith(LEGAL_ADDRESS_PREFIX)) {
                        legalAddress = readNextCellText(row, cell, formatter);
                    }
                }
                if (objectName.isEmpty()) {
                    objectName = extractObjectName(normalized);
                    if (objectName.isEmpty() && normalized.startsWith(OBJECT_NAME_PREFIX)) {
                        objectName = readNextCellText(row, cell, formatter);
                    }
                }
                if (objectAddress.isEmpty()) {
                    objectAddress = extractObjectAddress(normalized);
                    if (objectAddress.isEmpty() && normalized.startsWith(OBJECT_ADDRESS_PREFIX)) {
                        objectAddress = readNextCellText(row, cell, formatter);
                    }
                }
                if (!customer.isEmpty() && !dates.isEmpty() && !representative.isEmpty()
                        && !legalAddress.isEmpty() && !objectName.isEmpty() && !objectAddress.isEmpty()) {
                    return new MapHeaderData(customer, dates, representative, legalAddress, objectName, objectAddress);
                }
            }
        }
        return new MapHeaderData(customer, dates, representative, legalAddress, objectName, objectAddress);
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

    private static String extractLegalAddress(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(LEGAL_ADDRESS_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + LEGAL_ADDRESS_PREFIX.length()).trim();
    }

    private static String extractObjectName(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(OBJECT_NAME_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + OBJECT_NAME_PREFIX.length()).trim();
    }

    private static String extractObjectAddress(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(OBJECT_ADDRESS_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + OBJECT_ADDRESS_PREFIX.length()).trim();
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

    private static String resolveControlDate(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findControlDate(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findControlDate(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        Row row = sheet.getRow(6);
        if (row == null) {
            return "";
        }
        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (text.isEmpty()) {
                continue;
            }
            java.util.regex.Matcher matcher = CONTROL_DATE_PATTERN.matcher(text);
            if (matcher.find()) {
                return matcher.group();
            }
        }
        return "";
    }

    private static String resolveSpecialConditions(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            for (int idx = 0; idx < workbook.getNumberOfSheets(); idx++) {
                Sheet sheet = workbook.getSheetAt(idx);
                if (sheet == null) {
                    continue;
                }
                String normalized = normalizeText(sheet.getSheetName()).toLowerCase(Locale.ROOT);
                if (normalized.equals("эроа радона")) {
                    return "заказчик сообщил, что перед измерением ЭРОА радона " +
                            "здание выдерженно в течении более 12 часов при закрытых дверях и окнах";
                }
            }
            return "";
        } catch (Exception ex) {
            return "";
        }
    }

    private static String resolveMeasurementMethods(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findMeasurementMethods(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static java.util.List<InstrumentData> resolveMeasurementInstruments(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return java.util.Collections.emptyList();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return java.util.Collections.emptyList();
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findMeasurementInstruments(sheet);
        } catch (Exception ex) {
            return java.util.Collections.emptyList();
        }
    }

    private static java.util.List<InstrumentData> findMeasurementInstruments(Sheet sheet) {
        if (sheet == null) {
            return java.util.Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        int sectionRowIndex = -1;
        int headerRowIndex = -1;
        int nameColumn = -1;
        int serialColumn = -1;
        for (Row row : sheet) {
            for (Cell cell : row) {
                String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                if (normalized.equals(INSTRUMENTS_SECTION_HEADER)) {
                    if (sectionRowIndex < 0) {
                        sectionRowIndex = row.getRowNum();
                    }
                }
            }
        }

        if (sectionRowIndex >= 0) {
            for (int rowIndex = sectionRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                for (Cell cell : row) {
                    String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                    if (normalized.equals(INSTRUMENTS_NAME_HEADER)) {
                        nameColumn = cell.getColumnIndex();
                        headerRowIndex = rowIndex;
                        break;
                    }
                }
                if (headerRowIndex >= 0) {
                    break;
                }
            }
        }

        if (headerRowIndex < 0 || nameColumn < 0) {
            return java.util.Collections.emptyList();
        }

        Row headerRow = sheet.getRow(headerRowIndex);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                if (normalized.contains("завод")) {
                    serialColumn = cell.getColumnIndex();
                    break;
                }
            }
        }

        if (serialColumn < 0) {
            serialColumn = nameColumn + 19;
        }

        java.util.List<InstrumentData> instruments = new java.util.ArrayList<>();
        java.util.Set<String> seenSerials = new java.util.HashSet<>();
        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            String name = readMergedCellValue(sheet, rowIndex, nameColumn, formatter);
            String serial = readMergedCellValue(sheet, rowIndex, serialColumn, formatter);
            String normalizedName = normalizeText(name).toLowerCase(Locale.ROOT);
            if (normalizedName.equals(INSTRUMENTS_NAME_HEADER)) {
                continue;
            }
            if (name.isBlank() && serial.isBlank()) {
                break;
            }
            String serialKey = normalizeText(serial).toLowerCase(Locale.ROOT);
            if (!serialKey.isBlank() && !seenSerials.add(serialKey)) {
                continue;
            }
            instruments.add(new InstrumentData(name, serial));
        }
        return instruments;
    }

    private static String readMergedCellValue(Sheet sheet, int rowIndex, int colIndex, DataFormatter formatter) {
        if (sheet == null) {
            return "";
        }
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(rowIndex, colIndex)) {
                Row row = sheet.getRow(range.getFirstRow());
                if (row == null) {
                    return "";
                }
                Cell cell = row.getCell(range.getFirstColumn());
                return normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
            }
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
    }

    private static String findMeasurementMethods(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell);
                String normalized = normalizeText(rawText);
                if (!normalized.equals(METHODS_HEADER)) {
                    continue;
                }
                return collectMethodsBelow(sheet, row.getRowNum(), cell.getColumnIndex(), formatter);
            }
        }
        return "";
    }

    private static String collectMethodsBelow(Sheet sheet, int headerRow, int columnIndex, DataFormatter formatter) {
        java.util.List<String> methods = new java.util.ArrayList<>();
        for (int rowIndex = headerRow + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                break;
            }
            Cell cell = row.getCell(columnIndex);
            String raw = cell == null ? "" : formatter.formatCellValue(cell);
            String normalized = normalizeText(raw);
            if (normalized.isEmpty()) {
                break;
            }
            if (normalized.equals(METHODS_HEADER)) {
                continue;
            }
            methods.add(normalized);
        }
        return String.join("; ", methods);
    }

    private static boolean isGeneratorSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.equals("генератор")
                || normalized.equals("генератор (2)")
                || normalized.equals("генератор (2.0.)");
    }

    private static String findMeasurementPerformer(Sheet sheet, DataFormatter formatter) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell).trim();
                if (rawText.isEmpty()) {
                    continue;
                }
                String text = normalizeText(rawText);
                if (text.contains("Измерения проводил")) {
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

    private static String resolveProtocolNumber(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findProtocolNumber(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findProtocolNumber(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                int index = text.indexOf(PROTOCOL_PREFIX);
                if (index < 0) {
                    continue;
                }
                String tail = text.substring(index + PROTOCOL_PREFIX.length()).trim();
                tail = stripLeadingNumberMarker(tail);
                if (!tail.isEmpty()) {
                    return tail;
                }
                String nextText = readNextCellText(row, cell, formatter);
                return stripLeadingNumberMarker(nextText);
            }
        }
        return "";
    }

    private static String resolveContractText(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findContractText(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findContractText(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                int index = text.indexOf(BASIS_PREFIX);
                if (index < 0) {
                    continue;
                }
                String tail = text.substring(index + BASIS_PREFIX.length()).trim();
                if (!tail.isEmpty()) {
                    return tail;
                }
                return readNextCellText(row, cell, formatter);
            }
        }
        return "";
    }

    private static String normalizeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
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

    private static String stripLeadingNumberMarker(String value) {
        if (value == null) {
            return "";
        }
        String trimmed = value.trim();
        if (trimmed.startsWith("№")) {
            return trimmed.substring(1).trim();
        }
        return trimmed;
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

    private static void adjustRowHeightForMergedText(Sheet sheet,
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

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * Math.max(1, lines))));
    }

    private static void adjustRowHeightForInstrumentRow(Sheet sheet,
                                                        int rowIndex,
                                                        String name,
                                                        String serialNumber) {
        if (sheet == null) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double nameChars = totalColumnChars(sheet, 1, 19);
        double serialChars = totalColumnChars(sheet, 20, 29);
        int nameLines = estimateWrappedLines(name, nameChars);
        int serialLines = estimateWrappedLines(serialNumber, serialChars);
        int lines = Math.max(1, Math.max(nameLines, serialLines));

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * lines)));
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
    private static final String LEGAL_ADDRESS_PREFIX = "Юридический адрес заказчика:";
    private static final String OBJECT_NAME_PREFIX =
            "Наименование предприятия, организации, объекта, где производились измерения:";
    private static final String OBJECT_ADDRESS_PREFIX = "Адрес предприятия (объекта):";
    private static final String METHODS_HEADER =
            "Документы, устанавливающие правила и методы исследований (испытаний) и измерений";
    private static final String INSTRUMENTS_SECTION_HEADER = "сведения о средствах измерения:";
    private static final String INSTRUMENTS_NAME_HEADER = "наименование, тип средства измерения";
    private static final String PROTOCOL_PREFIX = "Протокол испытаний";
    private static final String BASIS_PREFIX = "Основание для измерений: договор";
    private static final java.util.regex.Pattern DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{2}\\.\\d{2}\\.\\d{4}\\b");
    private static final java.util.regex.Pattern CONTROL_DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{1,2}\\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s+\\d{4}\\b",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    private static final class MapHeaderData {
        private final String customerNameAndContacts;
        private final String measurementDates;
        private final String representative;
        private final String customerLegalAddress;
        private final String objectName;
        private final String objectAddress;

        private MapHeaderData(String customerNameAndContacts,
                              String measurementDates,
                              String representative,
                              String customerLegalAddress,
                              String objectName,
                              String objectAddress) {
            this.customerNameAndContacts = customerNameAndContacts;
            this.measurementDates = measurementDates;
            this.representative = representative;
            this.customerLegalAddress = customerLegalAddress;
            this.objectName = objectName;
            this.objectAddress = objectAddress;
        }
    }

    private static final class InstrumentData {
        private final String name;
        private final String serialNumber;

        private InstrumentData(String name, String serialNumber) {
            this.name = safe(name);
            this.serialNumber = safe(serialNumber);
        }
    }
}
