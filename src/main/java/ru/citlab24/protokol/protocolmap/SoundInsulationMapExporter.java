package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.util.Units;

import javax.imageio.ImageIO;
import java.awt.Dimension;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.IdentityHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class SoundInsulationMapExporter {
    private static final String PRIMARY_FOLDER_NAME = "Первичка Звукоизоляция";
    private static final String REGISTRATION_LABEL = "9. Регистрационный номер карты:";
    private static final String CUSTOMER_LABEL = "Наименование и контактные данные заявителя (заказчика):";
    private static final String MEASUREMENT_DATES_LABEL = "Дата проведения измерений:";
    private static final String PERFORMER_LABEL = "Протокол подготовил";
    private static final String REPRESENTATIVE_LABEL = "7. Измерения проводились в присутствии:";
    private static final String PROTOCOL_NUMBER_LABEL = "Протокол испытаний №";
    private static final String CONTRACT_LABEL = "6. Основание для измерений: договор";
    private static final String LEGAL_ADDRESS_LABEL = "2. Юридический адрес заказчика:";
    private static final String OBJECT_NAME_LABEL =
            "4. Наименование предприятия, организации, объекта, где производились измерения:";
    private static final String OBJECT_ADDRESS_LABEL = "5. Адрес предприятия (объекта):";
    private static final String OBJECT_ADDRESS_SECTION_LABEL = "5. Адрес предприятия";
    private static final String APPROVAL_LABEL = "УТВЕРЖДАЮ";
    private static final Pattern APPROVAL_DATE_PATTERN =
            Pattern.compile("\\b\\d{1,2}\\s+[А-Яа-я]+\\s+\\d{4}\\s*г?\\.?");
    private static final double SKETCH_HEIGHT_CM = 6.26;
    private static final double SKETCH_WIDTH_CM = 4.08;
    private static final int SKETCH_PADDING_PX = 5;

    private SoundInsulationMapExporter() {
    }

    public static File generateMap(File impactFile,
                                   File wallFile,
                                   File slabFile,
                                   File protocolFile,
                                   String workDeadline,
                                   String customerInn) throws IOException {
        File targetFile = PhysicalFactorsMapExporter.generateMap(impactFile, workDeadline, customerInn,
                PRIMARY_FOLDER_NAME);
        removeMicroclimateSheet(targetFile);
        removeVentilationSheet(targetFile);
        SoundInsulationProtocolData data = new SoundInsulationProtocolData();
        boolean hasProtocol = protocolFile != null && protocolFile.exists();
        if (hasProtocol) {
            try {
                data = extractProtocolData(protocolFile);
                applyProtocolData(targetFile, data);
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        updateSketchesSheet(targetFile, data.sketches);
        applySecondPageFontSize(targetFile, 10);
        addLnwAcSheet(targetFile, impactFile);
        return targetFile;
    }

    private static void removeMicroclimateSheet(File targetFile) throws IOException {
        if (targetFile == null || !targetFile.exists()) {
            return;
        }
        try (InputStream inputStream = new FileInputStream(targetFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            List<Integer> indicesToRemove = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                String name = workbook.getSheetName(i);
                if (name != null && name.startsWith("Микроклимат")) {
                    indicesToRemove.add(i);
                }
            }
            for (int i = indicesToRemove.size() - 1; i >= 0; i--) {
                workbook.removeSheetAt(indicesToRemove.get(i));
            }
            if (!indicesToRemove.isEmpty()) {
                try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                    workbook.write(outputStream);
                }
            }
        }
    }

    private static void removeVentilationSheet(File targetFile) throws IOException {
        if (targetFile == null || !targetFile.exists()) {
            return;
        }
        try (InputStream inputStream = new FileInputStream(targetFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            List<Integer> indicesToRemove = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                String name = workbook.getSheetName(i);
                if (name != null && name.equalsIgnoreCase("Вентиляция")) {
                    indicesToRemove.add(i);
                }
            }
            for (int i = indicesToRemove.size() - 1; i >= 0; i--) {
                workbook.removeSheetAt(indicesToRemove.get(i));
            }
            if (!indicesToRemove.isEmpty()) {
                try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                    workbook.write(outputStream);
                }
            }
        }
    }

    private static SoundInsulationProtocolData extractProtocolData(File protocolFile) throws IOException {
        if (protocolFile == null || !protocolFile.exists()) {
            return new SoundInsulationProtocolData();
        }
        if (!protocolFile.getName().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            return new SoundInsulationProtocolData();
        }
        try (FileInputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document);
            String registrationNumber = extractValueAfterLabel(lines, REGISTRATION_LABEL);
            String customer = extractValueAfterLabel(lines, CUSTOMER_LABEL);
            String measurementDates = extractValueAfterLabel(lines, MEASUREMENT_DATES_LABEL);
            String representative = extractValueAfterLabel(lines, REPRESENTATIVE_LABEL);
            String protocolNumber = extractValueAfterLabel(lines, PROTOCOL_NUMBER_LABEL);
            String contractText = extractValueAfterLabel(lines, CONTRACT_LABEL);
            String legalAddress = extractValueAfterLabel(lines, LEGAL_ADDRESS_LABEL);
            String objectName = extractBetweenLabels(lines, OBJECT_NAME_LABEL, OBJECT_ADDRESS_SECTION_LABEL);
            String objectAddress = extractValueAfterLabel(lines, OBJECT_ADDRESS_LABEL);
            String controlDate = extractApprovalDate(lines);
            String measurementPerformer = resolveMeasurementPerformer(document, lines);
            String measurementMethods = extractMeasurementMethods(document);
            List<InstrumentData> instruments = extractInstruments(document);
            List<String> roomNames = extractRoomNames(document);
            String objectDetails = extractObjectDetailsBlock(document);
            String constructiveSolutionsTable = extractTableTextAfterTitle(document, "13.1 Конструктивные решения:");
            String roomParametersTable = extractTableTextAfterTitle(document,
                    "16. Параметры помещений и испытываемой поверхности:");
            String areaBetweenRooms = extractLineContaining(lines, "Площадь испытываемой поверхности между помещениями");
            List<SketchEntry> sketches = extractSketches(document);
            String controlPerson = resolveControlPerson(measurementPerformer);
            return new SoundInsulationProtocolData(registrationNumber, customer, measurementDates,
                    measurementPerformer, representative, controlPerson, controlDate, protocolNumber, contractText,
                    legalAddress, objectName, objectAddress, measurementMethods, instruments, roomNames,
                    objectDetails, constructiveSolutionsTable, roomParametersTable, areaBetweenRooms, sketches);
        }
    }

    private static List<String> extractLines(XWPFDocument document) {
        List<String> lines = new ArrayList<>();
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                addParagraphLines(lines, paragraph.getText());
            } else if (element instanceof XWPFTable table) {
                addTableLines(lines, table);
            }
        }
        return lines;
    }

    private static void addParagraphLines(List<String> lines, String text) {
        if (text == null) {
            return;
        }
        String[] split = text.split("\\R");
        for (String line : split) {
            String normalized = normalizeSpace(line);
            if (!normalized.isBlank()) {
                lines.add(normalized);
            }
        }
    }

    private static void addTableLines(List<String> lines, XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            StringBuilder builder = new StringBuilder();
            for (XWPFTableCell cell : row.getTableCells()) {
                String cellText = normalizeSpace(cell.getText());
                if (cellText.isBlank()) {
                    continue;
                }
                if (builder.length() > 0) {
                    builder.append(' ');
                }
                builder.append(cellText);
            }
            String rowText = builder.toString().trim();
            if (!rowText.isBlank()) {
                addParagraphLines(lines, rowText);
            }
        }
    }

    private static String extractValueAfterLabel(List<String> lines, String label) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String labelNormalized = normalizeSpace(label).toLowerCase(Locale.ROOT);
        for (int i = 0; i < lines.size(); i++) {
            String line = normalizeSpace(lines.get(i));
            if (line.isBlank()) {
                continue;
            }
            String lower = line.toLowerCase(Locale.ROOT);
            int index = lower.indexOf(labelNormalized);
            if (index < 0) {
                continue;
            }
            String remainder = stripLeadingSeparators(line.substring(index + labelNormalized.length()));
            if (remainder.isBlank() && i + 1 < lines.size()) {
                remainder = normalizeSpace(lines.get(i + 1));
            }
            return remainder;
        }
        return "";
    }

    private static String extractBetweenLabels(List<String> lines, String startLabel, String endLabel) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String startNormalized = normalizeSpace(startLabel).toLowerCase(Locale.ROOT);
        String endNormalized = normalizeSpace(endLabel).toLowerCase(Locale.ROOT);
        for (int i = 0; i < lines.size(); i++) {
            String line = normalizeSpace(lines.get(i));
            String lower = line.toLowerCase(Locale.ROOT);
            int startIndex = lower.indexOf(startNormalized);
            if (startIndex < 0) {
                continue;
            }
            StringBuilder builder = new StringBuilder();
            String remainder = stripLeadingSeparators(line.substring(startIndex + startNormalized.length()));
            if (!remainder.isBlank()) {
                builder.append(remainder);
            }
            for (int j = i + 1; j < lines.size(); j++) {
                String nextLine = normalizeSpace(lines.get(j));
                if (nextLine.isBlank()) {
                    continue;
                }
                String nextLower = nextLine.toLowerCase(Locale.ROOT);
                int endIndex = nextLower.indexOf(endNormalized);
                if (endIndex >= 0) {
                    String beforeEnd = stripLeadingSeparators(nextLine.substring(0, endIndex));
                    if (!beforeEnd.isBlank()) {
                        appendWithSpace(builder, beforeEnd);
                    }
                    return builder.toString().trim();
                }
                appendWithSpace(builder, nextLine);
            }
            return builder.toString().trim();
        }
        return "";
    }

    private static String extractApprovalDate(List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        for (int i = 0; i < lines.size(); i++) {
            String line = normalizeSpace(lines.get(i));
            if (line.isBlank()) {
                continue;
            }
            if (!line.toUpperCase(Locale.ROOT).contains(APPROVAL_LABEL)) {
                continue;
            }
            String date = findDateInLine(line);
            if (!date.isBlank()) {
                return date;
            }
            if (i + 1 < lines.size()) {
                date = findDateInLine(lines.get(i + 1));
                if (!date.isBlank()) {
                    return date;
                }
            }
        }
        return "";
    }

    private static String findDateInLine(String line) {
        if (line == null || line.isBlank()) {
            return "";
        }
        Matcher matcher = APPROVAL_DATE_PATTERN.matcher(line);
        if (matcher.find()) {
            return matcher.group().trim();
        }
        return "";
    }

    private static String resolveMeasurementPerformer(XWPFDocument document, List<String> lines) {
        return "инженер Белов Д.А.";
    }

    private static String resolveMeasurementPerformerFromLines(List<String> lines, String defaultPerformer) {
        if (lines == null || lines.isEmpty()) {
            return defaultPerformer;
        }
        for (int i = 0; i < lines.size(); i++) {
            String line = normalizeSpace(lines.get(i));
            if (line.isBlank()) {
                continue;
            }
            String lower = line.toLowerCase(Locale.ROOT);
            if (!lower.contains(PERFORMER_LABEL.toLowerCase(Locale.ROOT))) {
                continue;
            }
            StringBuilder combined = new StringBuilder(line);
            if (i + 1 < lines.size()) {
                combined.append(' ').append(normalizeSpace(lines.get(i + 1)));
            }
            String combinedLower = combined.toString().toLowerCase(Locale.ROOT);
            if (combinedLower.contains("белов")) {
                return "инженер Белов Д.А.";
            }
            return defaultPerformer;
        }
        return defaultPerformer;
    }

    private static String resolveControlPerson(String measurementPerformer) {
        if (measurementPerformer == null) {
            return "Белов Д.А.";
        }
        String lower = measurementPerformer.toLowerCase(Locale.ROOT);
        if (lower.contains("белов")) {
            return "Тарновский М.О.";
        }
        return "Белов Д.А.";
    }

    private static void applyProtocolData(File targetFile, SoundInsulationProtocolData data) throws IOException {
        try (InputStream inputStream = new FileInputStream(targetFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheet("карта замеров");
            if (sheet == null) {
                return;
            }

            Font prefixFont = workbook.createFont();
            prefixFont.setFontName("Arial");
            prefixFont.setFontHeightInPoints((short) 14);
            prefixFont.setBold(true);

            Font valueFont = workbook.createFont();
            valueFont.setFontName("Arial");
            valueFont.setFontHeightInPoints((short) 12);
            valueFont.setBold(false);

            setCellText(sheet, 3, "КАРТА ЗАМЕРОВ № " + safe(data.registrationNumber));
            setMergedCellValueWithPrefix(sheet, 5, "1. Заказчик: ", data.customerNameAndContacts, prefixFont, valueFont);
            adjustRowHeightForMergedTextDoubling(sheet, 5, 0, 31,
                    "1. Заказчик: " + safe(data.customerNameAndContacts));
            setMergedCellValueWithPrefix(sheet, 7, "2. Дата замеров: ", data.measurementDates, prefixFont, valueFont);
            adjustRowHeightForMergedTextDoubling(sheet, 7, 0, 31,
                    "2. Дата замеров: " + safe(data.measurementDates));
            setMergedCellValueWithPrefix(sheet, 9, "3. Измерения провел, подпись: ",
                    data.measurementPerformer, prefixFont, valueFont);
            setMergedCellValueWithPrefix(sheet, 11, "4. Измерения проведены в присутствии представителя: ",
                    data.representative, prefixFont, valueFont);
            adjustRowHeightForMergedTextDoubling(sheet, 11, 0, 31,
                    "4. Измерения проведены в присутствии представителя: " + safe(data.representative));
            setMergedCellValueWithPrefix(sheet, 15, "5. Контроль ведения записей осуществлен: ",
                    data.controlPerson, prefixFont, valueFont);
            setMergedCellValueWithPrefix(sheet, 19, "Дата контроля: ", data.controlDate, prefixFont, valueFont);

            setCellText(sheet, 21, "1. Номер протокола " + safe(data.protocolNumber));
            setCellText(sheet, 22, "2. Договор " + safe(data.contractText));
            setCellText(sheet, 23, "3. Наименование и контактные данные Заказчика: "
                    + safe(data.customerNameAndContacts));
            setCellText(sheet, 24, "Юридический адрес заказчика: " + safe(data.customerLegalAddress));

            String objectNameText = "4. Наименование объекта: " + safe(data.objectName);
            setCellText(sheet, 25, objectNameText);
            adjustRowHeightForMergedText(sheet, 25, 0, 31, objectNameText);

            String objectAddressText = "Адрес объекта " + safe(data.objectAddress);
            setCellText(sheet, 26, objectAddressText);
            adjustRowHeightForMergedText(sheet, 26, 0, 31, objectAddressText);

            applyHeaderCenter(sheet, data.registrationNumber);
            updateSpecialConditions(sheet);
            updateMeasurementMethods(sheet, data.measurementMethods);
            updateInstrumentsTable(sheet, data.instruments);
            updateSketchSection(sheet, data);

            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void setCellText(Sheet sheet, int rowIndex, String text) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            cell = row.createCell(0);
        }
        cell.setCellValue(text);
    }

    private static void applyHeaderCenter(Sheet sheet, String registrationNumber) {
        if (sheet == null) {
            return;
        }
        Header header = sheet.getHeader();
        String font = "&\"Arial\"&12";
        header.setCenter(font + "Карта замеров № " + safe(registrationNumber) + "\nФ8 РИ ИЛ 2-2023");
    }

    private static void applySecondPageFontSize(File targetFile, int fontSize) throws IOException {
        if (targetFile == null || !targetFile.exists()) {
            return;
        }
        try (InputStream inputStream = new FileInputStream(targetFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            Sheet sheet = workbook.getSheet("карта замеров");
            if (sheet == null) {
                return;
            }
            int startRow = 21;
            int lastRow = sheet.getLastRowNum();
            Map<CellStyle, CellStyle> styleCache = new IdentityHashMap<>();
            Map<Integer, Font> fontCache = new java.util.HashMap<>();
            for (int rowIndex = startRow; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                for (Cell cell : row) {
                    CellStyle style = cell.getCellStyle();
                    if (style == null) {
                        continue;
                    }
                    CellStyle updatedStyle = styleCache.get(style);
                    if (updatedStyle == null) {
                        updatedStyle = workbook.createCellStyle();
                        updatedStyle.cloneStyleFrom(style);
                        int fontIndex = style.getFontIndexAsInt();
                        Font originalFont = workbook.getFontAt(fontIndex);
                        Font resizedFont = fontCache.get(fontIndex);
                        if (resizedFont == null) {
                            resizedFont = cloneFontWithSize(workbook, originalFont, fontSize);
                            fontCache.put(fontIndex, resizedFont);
                        }
                        updatedStyle.setFont(resizedFont);
                        styleCache.put(style, updatedStyle);
                    }
                    cell.setCellStyle(updatedStyle);
                }
            }
            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void addLnwAcSheet(File targetFile, File impactFile) throws IOException {
        if (targetFile == null || !targetFile.exists() || impactFile == null || !impactFile.exists()) {
            return;
        }
        try (InputStream targetInput = new FileInputStream(targetFile);
             Workbook targetWorkbook = WorkbookFactory.create(targetInput);
             InputStream impactInput = new FileInputStream(impactFile);
             Workbook impactWorkbook = WorkbookFactory.create(impactInput)) {
            int existingIndex = targetWorkbook.getSheetIndex("Lnw AC");
            if (existingIndex >= 0) {
                targetWorkbook.removeSheetAt(existingIndex);
            }
            Sheet targetSheet = targetWorkbook.createSheet("Lnw AC");
            applyLnwAcColumnWidths(targetSheet);

            LnwAcSourceData sourceData = resolveLnwAcSourceData(impactWorkbook);
            String headerText = buildLnwAcHeaderText(sourceData.firstRoomWord, sourceData.secondRoomWord);
            createLnwAcHeaderRow(targetWorkbook, targetSheet, headerText);

            if (sourceData.sourceSheet != null && sourceData.startRow >= 0 && sourceData.endRow >= sourceData.startRow) {
                copyLnwAcRows(sourceData, targetWorkbook, targetSheet, 1);
            }

            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                targetWorkbook.write(outputStream);
            }
        }
    }

    private static void applyLnwAcColumnWidths(Sheet sheet) {
        int[] widthsPx = new int[17];
        widthsPx[0] = 173;
        for (int i = 1; i < widthsPx.length; i++) {
            widthsPx[i] = 54;
        }
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
        }
    }

    private static void createLnwAcHeaderRow(Workbook workbook, Sheet sheet, String headerText) {
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(headerText);

        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        cell.setCellStyle(style);

        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 16));
    }

    private static void copyLnwAcRows(LnwAcSourceData sourceData,
                                      Workbook targetWorkbook,
                                      Sheet targetSheet,
                                      int targetStartRow) {
        DataFormatter formatter = new DataFormatter();
        Map<CellStyle, CellStyle> styleCache = new IdentityHashMap<>();
        int rowOffset = targetStartRow - sourceData.startRow;
        for (int rowIndex = sourceData.startRow; rowIndex <= sourceData.endRow; rowIndex++) {
            Row sourceRow = sourceData.sourceSheet.getRow(rowIndex);
            int targetRowIndex = rowIndex + rowOffset;
            Row targetRow = targetSheet.getRow(targetRowIndex);
            if (targetRow == null) {
                targetRow = targetSheet.createRow(targetRowIndex);
            }
            if (sourceRow != null) {
                targetRow.setHeight(sourceRow.getHeight());
            }
            for (int col = 0; col <= 16; col++) {
                Cell sourceCell = sourceRow != null ? sourceRow.getCell(col) : null;
                if (sourceCell == null) {
                    continue;
                }
                Cell targetCell = targetRow.getCell(col);
                if (targetCell == null) {
                    targetCell = targetRow.createCell(col);
                }
                copyCellValue(sourceCell, targetCell, formatter);
                CellStyle sourceStyle = sourceCell.getCellStyle();
                if (sourceStyle != null) {
                    CellStyle clonedStyle = styleCache.get(sourceStyle);
                    if (clonedStyle == null) {
                        clonedStyle = targetWorkbook.createCellStyle();
                        clonedStyle.cloneStyleFrom(sourceStyle);
                        styleCache.put(sourceStyle, clonedStyle);
                    }
                    targetCell.setCellStyle(clonedStyle);
                }
            }
        }
        copyMergedRegions(sourceData, targetSheet, rowOffset);
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell, DataFormatter formatter) {
        if (sourceCell == null || targetCell == null) {
            return;
        }
        switch (sourceCell.getCellType()) {
            case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
            case NUMERIC -> targetCell.setCellValue(sourceCell.getNumericCellValue());
            case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
            case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            case BLANK -> targetCell.setBlank();
            default -> targetCell.setCellValue(formatter.formatCellValue(sourceCell));
        }
    }

    private static void copyMergedRegions(LnwAcSourceData sourceData, Sheet targetSheet, int rowOffset) {
        int count = sourceData.sourceSheet.getNumMergedRegions();
        for (int index = 0; index < count; index++) {
            CellRangeAddress region = sourceData.sourceSheet.getMergedRegion(index);
            if (region.getFirstRow() < sourceData.startRow || region.getLastRow() > sourceData.endRow) {
                continue;
            }
            if (region.getFirstColumn() > 16 || region.getLastColumn() > 16) {
                continue;
            }
            CellRangeAddress shifted = new CellRangeAddress(
                    region.getFirstRow() + rowOffset,
                    region.getLastRow() + rowOffset,
                    region.getFirstColumn(),
                    region.getLastColumn()
            );
            targetSheet.addMergedRegion(shifted);
        }
    }

    private static String buildLnwAcHeaderText(String firstRoomWord, String secondRoomWord) {
        String base = "7.8.1 Результаты измерения приведенного уровня ударного шума, "
                + "индекса приведенного уровня ударного шума для перекрытия между помещениями";
        String first = safe(firstRoomWord).trim();
        String second = safe(secondRoomWord).trim();
        if (first.isBlank() && second.isBlank()) {
            return base;
        }
        if (first.isBlank()) {
            return base + " " + second;
        }
        if (second.isBlank()) {
            return base + " " + first;
        }
        return base + " " + first + " и " + second;
    }

    private static LnwAcSourceData resolveLnwAcSourceData(Workbook impactWorkbook) {
        String firstRoomWord = "";
        String secondRoomWord = "";
        DataFormatter formatter = new DataFormatter();
        List<String> pnuWords = new ArrayList<>();
        int sheetCount = impactWorkbook.getNumberOfSheets();
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = impactWorkbook.getSheetAt(sheetIndex);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = normalizeSpace(formatter.formatCellValue(cell));
                    if (text.equalsIgnoreCase("ПНУ:")) {
                        Cell rightCell = row.getCell(cell.getColumnIndex() + 1);
                        String rightText = normalizeSpace(formatter.formatCellValue(rightCell));
                        if (!rightText.isBlank()) {
                            String firstWord = rightText.split("\\s+")[0];
                            pnuWords.add(firstWord);
                            if (pnuWords.size() >= 2) {
                                break;
                            }
                        }
                    }
                }
                if (pnuWords.size() >= 2) {
                    break;
                }
            }
            if (pnuWords.size() >= 2) {
                break;
            }
        }
        if (!pnuWords.isEmpty()) {
            firstRoomWord = pnuWords.get(0);
        }
        if (pnuWords.size() > 1) {
            secondRoomWord = pnuWords.get(1);
        }

        LnwAcSourceData sourceData = new LnwAcSourceData();
        sourceData.firstRoomWord = firstRoomWord;
        sourceData.secondRoomWord = secondRoomWord;

        String startNeedle = "для карты";
        String stopNeedle = "Текст цветом переписывать только с протоколов ФГИС";
        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = impactWorkbook.getSheetAt(sheetIndex);
            int startRow = -1;
            int endRow = -1;
            for (Row row : sheet) {
                if (startRow < 0) {
                    if (rowContainsText(row, startNeedle, formatter)) {
                        startRow = row.getRowNum() + 1;
                        continue;
                    }
                } else if (rowContainsText(row, stopNeedle, formatter)) {
                    endRow = row.getRowNum() - 1;
                    break;
                }
            }
            if (startRow >= 0) {
                if (endRow < startRow) {
                    endRow = sheet.getLastRowNum();
                }
                sourceData.sourceSheet = sheet;
                sourceData.startRow = startRow;
                sourceData.endRow = endRow;
                break;
            }
        }
        return sourceData;
    }

    private static boolean rowContainsText(Row row, String needle, DataFormatter formatter) {
        if (row == null) {
            return false;
        }
        String needleLower = needle.toLowerCase(Locale.ROOT);
        for (Cell cell : row) {
            String text = normalizeSpace(formatter.formatCellValue(cell));
            if (text.toLowerCase(Locale.ROOT).contains(needleLower)) {
                return true;
            }
        }
        return false;
    }

    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static Font cloneFontWithSize(Workbook workbook, Font source, int fontSize) {
        Font font = workbook.createFont();
        font.setFontName(source.getFontName());
        font.setBold(source.getBold());
        font.setItalic(source.getItalic());
        font.setUnderline(source.getUnderline());
        font.setColor(source.getColor());
        font.setCharSet(source.getCharSet());
        font.setStrikeout(source.getStrikeout());
        font.setTypeOffset(source.getTypeOffset());
        font.setFontHeightInPoints((short) fontSize);
        return font;
    }

    private static void updateSpecialConditions(Sheet sheet) {
        int rowIndex = findRowIndexByPrefix(sheet, "5.1. Особые условия:");
        if (rowIndex < 0) {
            return;
        }
        String text = "5.1. Особые условия: -";
        setCellText(sheet, rowIndex, text);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, text);
    }

    private static void updateMeasurementMethods(Sheet sheet, String measurementMethods) {
        int rowIndex = findRowIndexByPrefix(sheet, "5.2. Методы измерения");
        if (rowIndex < 0) {
            return;
        }
        String text = "5.2. Методы измерения " + safe(measurementMethods);
        setCellText(sheet, rowIndex, text);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, text);
    }

    private static void updateInstrumentsTable(Sheet sheet, List<InstrumentData> instruments) {
        int labelRow = findRowIndexByPrefix(sheet, "5.3. Приборы для измерения");
        if (labelRow < 0) {
            return;
        }
        applyWrapStyleToRange(sheet, labelRow, 0, 31);
        adjustRowHeightForMergedText(sheet, labelRow, 0, 31, readRowText(sheet, labelRow));
        int headerRow = labelRow + 1;
        int dataStartRow = headerRow + 1;
        int sketchRow = findRowIndexByPrefix(sheet, "6. Эскиз");
        int lastRow = sheet.getLastRowNum();
        int endBoundary = sketchRow > 0 ? sketchRow : lastRow;

        removeRowBreaks(sheet, labelRow, endBoundary);

        int lastInstrumentRow = dataStartRow - 1;
        for (int rowIndex = dataStartRow; rowIndex < endBoundary; rowIndex++) {
            if (isInstrumentRowEmpty(sheet, rowIndex)) {
                break;
            }
            lastInstrumentRow = rowIndex;
        }
        int existingRows = Math.max(0, lastInstrumentRow - dataStartRow + 1);
        int neededRows = instruments == null ? 0 : instruments.size();
        if (neededRows > existingRows) {
            int delta = neededRows - existingRows;
            sheet.shiftRows(endBoundary, lastRow, delta);
        }

        int totalRows = Math.max(existingRows, neededRows);
        for (int index = 0; index < totalRows; index++) {
            int rowIndex = dataStartRow + index;
            InstrumentData instrument = index < neededRows ? instruments.get(index) : null;
            updateInstrumentRow(sheet, rowIndex, instrument);
        }
    }

    private static void updateInstrumentRow(Sheet sheet, int rowIndex, InstrumentData instrument) {
        Row templateRow = sheet.getRow(rowIndex);
        if (templateRow == null) {
            templateRow = sheet.createRow(rowIndex);
        }
        Row fallbackRow = rowIndex > 0 ? sheet.getRow(rowIndex - 1) : null;
        CellStyle templateNameStyle = resolveCellStyle(templateRow, 1, resolveCellStyle(fallbackRow, 1, null));
        CellStyle templateSerialStyle = resolveCellStyle(templateRow, 20, resolveCellStyle(fallbackRow, 20, templateNameStyle));
        CellStyle templateCheckboxStyle = resolveCellStyle(templateRow, 30,
                resolveCellStyle(fallbackRow, 30, templateSerialStyle));

        ensureMergedRegion(sheet, rowIndex, 1, 19);
        ensureMergedRegion(sheet, rowIndex, 20, 29);

        for (int col = 1; col <= 29; col++) {
            Cell cell = templateRow.getCell(col);
            if (cell == null) {
                cell = templateRow.createCell(col);
            }
            if (col <= 19) {
                if (templateNameStyle != null) {
                    cell.setCellStyle(templateNameStyle);
                }
                if (col == 1) {
                    cell.setCellValue(instrument == null ? "" : safe(instrument.name));
                }
            } else {
                if (templateSerialStyle != null) {
                    cell.setCellStyle(templateSerialStyle);
                }
                if (col == 20) {
                    cell.setCellValue(instrument == null ? "" : safe(instrument.serialNumber));
                }
            }
        }

        Cell checkboxCell = templateRow.getCell(30);
        if (checkboxCell == null) {
            checkboxCell = templateRow.createCell(30);
        }
        if (templateCheckboxStyle != null) {
            checkboxCell.setCellStyle(templateCheckboxStyle);
        }
        checkboxCell.setCellValue(instrument == null ? "" : "☑");

        String name = instrument == null ? "" : safe(instrument.name);
        String serial = instrument == null ? "" : safe(instrument.serialNumber);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, name + " " + serial);
    }

    private static CellStyle resolveCellStyle(Row row, int col, CellStyle fallback) {
        if (row == null) {
            return fallback;
        }
        Cell cell = row.getCell(col);
        if (cell == null) {
            return fallback;
        }
        CellStyle style = cell.getCellStyle();
        return style == null ? fallback : style;
    }

    private static boolean isInstrumentRowEmpty(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return true;
        }
        DataFormatter formatter = new DataFormatter();
        String name = formatCellValue(formatter, row.getCell(1));
        String serial = formatCellValue(formatter, row.getCell(20));
        return name.isEmpty() && serial.isEmpty();
    }

    private static void ensureMergedRegion(Sheet sheet, int rowIndex, int firstCol, int lastCol) {
        for (int idx = 0; idx < sheet.getNumMergedRegions(); idx++) {
            if (sheet.getMergedRegion(idx).getFirstRow() == rowIndex
                    && sheet.getMergedRegion(idx).getLastRow() == rowIndex
                    && sheet.getMergedRegion(idx).getFirstColumn() == firstCol
                    && sheet.getMergedRegion(idx).getLastColumn() == lastCol) {
                return;
            }
        }
        sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol));
    }

    private static void updateSketchSection(Sheet sheet, SoundInsulationProtocolData data) {
        int sketchRow = findRowIndexByPrefix(sheet, "6. Эскиз");
        if (sketchRow < 0) {
            return;
        }
        if (sketchRow > 0) {
            sheet.removeRowBreak(sketchRow - 1);
        }
        int rowIndex = sketchRow;
        String calibrationText = "5.4. калибровочный уровень: 94 дБ\n" +
                "показания шумомера перед измерениями на частоте 1 кГц:\n" +
                "показания шумомера после измерений на частоте 1 кГц:";
        rowIndex = writeMergedRow(sheet, rowIndex, calibrationText);

        String meteorologyText = buildMeteorologyText(data.roomNames);
        if (!meteorologyText.isBlank()) {
            rowIndex = writeMergedRow(sheet, rowIndex, meteorologyText);
        }

        String measurementsText = "7.8. Измерения звукоизоляции ограждающих конструкций:\n" +
                "В соответствии с ГОСТ Р ИСО 3382-2-2013: точность метода - технический; " +
                "метод оценки кривых спада - расчет методом наименьших квадратов; " +
                "используемый метод усреднения результатов в каждой позиции - определением времени реверберации " +
                "для каждой из всех кривых спада и расчетом их среднего значения; " +
                "метод усреднения результатов по всем позициям - арифметическим усреднением времени реверберации; " +
                "пространственное среднее получают как среднее отдельных времен реверберации для всех независимых " +
                "измерительных конфигураций; применен звуковой сигнал - розовый шум.";
        rowIndex = writeMergedRow(sheet, rowIndex, measurementsText);

        rowIndex = writeMergedRow(sheet, rowIndex, "");

        String peopleText = "Количество людей присутствующих в помещениях при испытаниях - ___ человек\n" +
                "Объект испытаний – внутренние ограждающие конструкции помещений";
        rowIndex = writeMergedRow(sheet, rowIndex, peopleText);

        rowIndex = writeMergedRow(sheet, rowIndex, "");

        if (!safe(data.objectDetails).isBlank()) {
            rowIndex = writeMergedRowsForLines(sheet, rowIndex, data.objectDetails);
        }

        rowIndex = writeMergedRowWithBorders(sheet, rowIndex, "Конструктивные решения:");
        rowIndex = writeConstructiveSolutionsTable(sheet, rowIndex, data.constructiveSolutionsTable);

        rowIndex = writeMergedRow(sheet, rowIndex, "");
        rowIndex = writeMergedRow(sheet, rowIndex, "Параметры помещений:");
        rowIndex = writeRoomParametersTable(sheet, rowIndex, data.roomParametersTable);

        rowIndex = writeMergedRow(sheet, rowIndex, "");
        String areaLine = safe(data.areaBetweenRooms);
        if (!areaLine.isBlank()) {
            rowIndex = writeMergedRow(sheet, rowIndex, "Площадь испытываемой поверхности между помещениями: " +
                    areaLine.replaceFirst("(?i)^Площадь испытываемой поверхности между помещениями:?\\s*", ""));
        }
    }

    private static void updateSketchesSheet(File targetFile, List<SketchEntry> sketches) throws IOException {
        if (targetFile == null || !targetFile.exists()) {
            return;
        }
        try (InputStream inputStream = new FileInputStream(targetFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            int existingIndex = workbook.getSheetIndex("Эскизы");
            if (existingIndex >= 0) {
                workbook.removeSheetAt(existingIndex);
            }
            Sheet sheet = workbook.createSheet("Эскизы");
            applySketchesSheetDefaults(workbook, sheet);
            createSketchesHeader(sheet);
            fillSketches(sheet, sketches);
            try (FileOutputStream outputStream = new FileOutputStream(targetFile)) {
                workbook.write(outputStream);
            }
        }
    }

    private static void applySketchesSheetDefaults(Workbook workbook, Sheet sheet) {
        if (workbook == null || sheet == null) {
            return;
        }
        sheet.setColumnWidth(0, 40 * 256);
        sheet.setColumnWidth(1, 40 * 256);
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);
        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        for (int col = 0; col <= 1; col++) {
            sheet.setDefaultColumnStyle(col, baseStyle);
        }
    }

    private static void createSketchesHeader(Sheet sheet) {
        Row headerRow = sheet.createRow(0);
        Cell cell = headerRow.createCell(0);
        cell.setCellValue("Эскизы");
        Workbook workbook = sheet.getWorkbook();
        if (workbook != null) {
            Font font = workbook.createFont();
            font.setBold(true);
            font.setFontName("Arial");
            font.setFontHeightInPoints((short) 12);
            CellStyle style = workbook.createCellStyle();
            style.setFont(font);
            cell.setCellStyle(style);
        }
        ensureMergedRegion(sheet, 0, 0, 1);
    }

    private static void fillSketches(Sheet sheet, List<SketchEntry> sketches) {
        if (sheet == null) {
            return;
        }
        int rowIndex = 2;
        if (sketches == null || sketches.isEmpty()) {
            return;
        }

        Workbook workbook = sheet.getWorkbook();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper helper = workbook.getCreationHelper();

        for (SketchEntry sketch : sketches) {
            if (sketch == null || sketch.imageData == null) {
                continue;
            }

            int imageRowIndex = rowIndex;
            Row imageRow = sheet.createRow(imageRowIndex);

            double targetWidthPx = cmToPixels(SKETCH_WIDTH_CM);
            double targetHeightPx = cmToPixels(SKETCH_HEIGHT_CM);

            int rowHeightPx = (int) Math.ceil(targetHeightPx + (SKETCH_PADDING_PX * 2.0));
            imageRow.setHeightInPoints(pixelsToPoints(rowHeightPx));

            // эскиз занимает A..B (0..1) на текущей строке
            ensureMergedRegion(sheet, imageRowIndex, 0, 1);
            applyWrapStyleToRange(sheet, imageRowIndex, 0, 1);

            int pictureIndex = workbook.addPicture(sketch.imageData, sketch.pictureType);

            ClientAnchor anchor = helper.createClientAnchor();
            anchor.setCol1(0);
            anchor.setRow1(imageRowIndex);

            // ВАЖНО: без col2/row2 у POI остаются нули, и при row1>0 получается row2<row1 => target size < 0
            anchor.setCol2(2);               // 0..1 (A..B) => col2 = 2
            anchor.setRow2(imageRowIndex + 1);

            anchor.setDx1(Units.pixelToEMU(SKETCH_PADDING_PX));
            anchor.setDy1(Units.pixelToEMU(SKETCH_PADDING_PX));

            // (не обязательно, но полезно)
            try {
                anchor.setAnchorType(ClientAnchor.AnchorType.MOVE_AND_RESIZE);
            } catch (Exception ignore) {
                // если AnchorType недоступен в конкретной реализации — просто игнорируем
            }

            org.apache.poi.ss.usermodel.Picture picture = drawing.createPicture(anchor, pictureIndex);
            resizePictureToTarget(picture, sketch.imageData, targetWidthPx, targetHeightPx);

            rowIndex++;

            String caption = safe(sketch.caption);
            if (!caption.isBlank()) {
                Row captionRow = sheet.createRow(rowIndex);
                Cell captionCell = captionRow.createCell(0);
                captionCell.setCellValue(caption);

                ensureMergedRegion(sheet, rowIndex, 0, 1);
                applyWrapStyleToRange(sheet, rowIndex, 0, 1);
                adjustRowHeightForMergedText(sheet, rowIndex, 0, 1, caption);

                rowIndex++;
            }
        }
    }

    private static int writeMergedRow(Sheet sheet, int rowIndex, String text) {
        setCellText(sheet, rowIndex, text);
        ensureMergedRegion(sheet, rowIndex, 0, 31);
        applyWrapStyleToRange(sheet, rowIndex, 0, 31);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, text);
        return rowIndex + 1;
    }

    private static int writeMergedRowWithBorders(Sheet sheet, int rowIndex, String text) {
        int nextRow = writeMergedRow(sheet, rowIndex, text);
        if (text != null && !text.isBlank()) {
            applyBorderToMergedRegion(sheet, rowIndex, 0, 31);
        }
        return nextRow;
    }

    private static int writeMergedRowsForLines(Sheet sheet, int rowIndex, String text) {
        if (text == null || text.isBlank()) {
            return rowIndex;
        }
        String[] lines = text.split("\\R", -1);
        int currentRow = rowIndex;
        for (String line : lines) {
            currentRow = writeMergedRow(sheet, currentRow, line);
        }
        return currentRow;
    }

    private static String buildMeteorologyText(List<String> roomNames) {
        if (roomNames == null || roomNames.isEmpty()) {
            return "";
        }
        StringBuilder builder = new StringBuilder("7.1. Метеорологические факторы атмосферного воздуха:");
        for (String room : roomNames) {
            String roomName = normalizeSpace(room);
            if (roomName.isBlank()) {
                continue;
            }
            builder.append("\nпомещение ").append(roomName)
                    .append(": Температура, __________ºС Относительная влажность, __________% ")
                    .append("Давление,__________мм рт. ст., скорость движения воздуха ______________м/с.");
        }
        return builder.toString();
    }

    private static int findRowIndexByPrefix(Sheet sheet, String prefix) {
        if (sheet == null) {
            return -1;
        }
        String normalizedPrefix = normalizeSpace(prefix).toLowerCase(Locale.ROOT);
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            Cell cell = row.getCell(0);
            if (cell == null) {
                continue;
            }
            String text = normalizeSpace(formatter.formatCellValue(cell));
            if (text.toLowerCase(Locale.ROOT).startsWith(normalizedPrefix)) {
                return row.getRowNum();
            }
        }
        return -1;
    }

    private static String formatCellValue(DataFormatter formatter, Cell cell) {
        if (formatter == null || cell == null) {
            return "";
        }
        return formatter.formatCellValue(cell).trim();
    }

    private static void setMergedCellValueWithPrefix(Sheet sheet,
                                                     int rowIndex,
                                                     String prefix,
                                                     String value,
                                                     Font prefixFont,
                                                     Font valueFont) {
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
                sheet.getWorkbook().getCreationHelper().createRichTextString(text);
        int prefixLength = prefix == null ? 0 : prefix.length();
        int textLength = text.length();
        if (textLength > 0) {
            richText.applyFont(0, Math.min(prefixLength, textLength), prefixFont);
            if (textLength > prefixLength) {
                richText.applyFont(prefixLength, textLength, valueFont);
            }
        }
        cell.setCellValue(richText);
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

    private static String readRowText(Sheet sheet, int rowIndex) {
        if (sheet == null) {
            return "";
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        return formatCellValue(formatter, row.getCell(0));
    }

    private static void applyWrapStyleToRange(Sheet sheet, int rowIndex, int firstCol, int lastCol) {
        if (sheet == null) {
            return;
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Map<CellStyle, CellStyle> cache = new IdentityHashMap<>();
        for (int col = firstCol; col <= lastCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            CellStyle baseStyle = cell.getCellStyle();
            if (baseStyle != null && baseStyle.getWrapText()) {
                continue;
            }
            CellStyle wrappedStyle = cache.get(baseStyle);
            if (wrappedStyle == null) {
                Workbook workbook = sheet.getWorkbook();
                wrappedStyle = workbook.createCellStyle();
                if (baseStyle != null) {
                    wrappedStyle.cloneStyleFrom(baseStyle);
                }
                wrappedStyle.setWrapText(true);
                cache.put(baseStyle, wrappedStyle);
            }
            cell.setCellStyle(wrappedStyle);
        }
    }

    private static void removeRowBreaks(Sheet sheet, int startRow, int endRow) {
        if (sheet == null) {
            return;
        }
        int from = Math.max(0, startRow);
        int to = Math.max(from, endRow);
        for (int rowIndex = from; rowIndex <= to; rowIndex++) {
            if (sheet.isRowBroken(rowIndex)) {
                sheet.removeRowBreak(rowIndex);
            }
        }
    }

    private static double totalColumnChars(Sheet sheet, int firstCol, int lastCol) {
        double total = 0;
        for (int col = firstCol; col <= lastCol; col++) {
            total += sheet.getColumnWidth(col) / 256.0;
        }
        return total;
    }

    private static int estimateWrappedLines(String text, double charsPerLine) {
        if (text == null || text.isBlank() || charsPerLine <= 0) {
            return 1;
        }
        int lines = 0;
        String[] paragraphs = text.split("\\R");
        for (String paragraph : paragraphs) {
            String trimmed = paragraph.trim();
            if (trimmed.isEmpty()) {
                lines++;
                continue;
            }
            int length = trimmed.length();
            lines += (int) Math.ceil(length / charsPerLine);
        }
        return Math.max(1, lines);
    }

    private static float pointsToPixels(float points) {
        return points * 96f / 72f;
    }

    private static float pixelsToPoints(int pixels) {
        return pixels * 72f / 96f;
    }

    private static double cmToPixels(double cm) {
        return cm / 2.54 * 96.0;
    }

    private static void resizePictureToTarget(org.apache.poi.ss.usermodel.Picture picture,
                                              byte[] imageData,
                                              double targetWidthPx,
                                              double targetHeightPx) {
        if (picture == null || imageData == null) {
            return;
        }
        Dimension dimension = readImageDimension(imageData);
        if (dimension == null || dimension.getWidth() <= 0 || dimension.getHeight() <= 0) {
            picture.resize();
            return;
        }
        double scaleX = targetWidthPx / dimension.getWidth();
        double scaleY = targetHeightPx / dimension.getHeight();
        picture.resize(scaleX, scaleY);
    }

    private static Dimension readImageDimension(byte[] imageData) {
        try (ByteArrayInputStream stream = new ByteArrayInputStream(imageData)) {
            java.awt.image.BufferedImage image = ImageIO.read(stream);
            if (image == null) {
                return null;
            }
            return new Dimension(image.getWidth(), image.getHeight());
        } catch (IOException ex) {
            return null;
        }
    }

    private static String safe(String value) {
        return value == null ? "" : value;
    }

    private static void appendWithSpace(StringBuilder builder, String value) {
        if (builder.length() > 0) {
            builder.append(' ');
        }
        builder.append(value);
    }

    private static String stripLeadingSeparators(String value) {
        if (value == null) {
            return "";
        }
        String trimmed = value.trim();
        while (trimmed.startsWith(":") || trimmed.startsWith("-") || trimmed.startsWith("—")) {
            trimmed = trimmed.substring(1).trim();
        }
        return trimmed;
    }

    private static String normalizeSpace(String value) {
        if (value == null) {
            return "";
        }
        return value.replaceAll("\\s+", " ").trim();
    }

    private static String extractMeasurementMethods(XWPFDocument document) {
        XWPFTable table = findTableAfterTitle(document,
                "12. Сведения о нормативных документах (НД), регламентирующих значения показателей и НД " +
                        "на методы (методики) измерений:");
        if (table == null) {
            return "";
        }
        List<String> values = extractColumnValues(table, 2, 1);
        return String.join("; ", values);
    }

    private static List<InstrumentData> extractInstruments(XWPFDocument document) {
        List<InstrumentData> instruments = new ArrayList<>();
        XWPFTable instrumentsTable = findTableAfterTitle(document, "10. Сведения о средствах измерения:");
        if (instrumentsTable != null) {
            instruments.addAll(extractInstrumentRows(instrumentsTable, 1, 2));
        }
        XWPFTable equipmentTable = findTableAfterTitle(document, "11. Сведения об испытательном оборудовании:");
        if (equipmentTable != null) {
            instruments.addAll(extractInstrumentRows(equipmentTable, 0, 1));
        }
        return instruments;
    }

    private static List<String> extractRoomNames(XWPFDocument document) {
        XWPFTable table = findTableAfterTitle(document, "16. Параметры помещений и испытываемой поверхности:");
        if (table == null) {
            return new ArrayList<>();
        }
        return extractColumnValues(table, 0, 1);
    }

    private static String extractObjectDetailsBlock(XWPFDocument document) {
        if (document == null) {
            return "";
        }
        String start = "Объект испытаний – внутренние ограждающие конструкции помещений, их монтаж осуществлен " +
                "заказчиком согласно требованиям технической документации.";
        String end = "13.1 Конструктивные решения:";
        StringBuilder builder = new StringBuilder();
        boolean capture = false;
        for (IBodyElement element : document.getBodyElements()) {
            if (!(element instanceof XWPFParagraph paragraph)) {
                continue;
            }
            String text = normalizeSpace(paragraph.getText());
            if (text.isBlank()) {
                continue;
            }
            if (!capture && text.contains(start)) {
                capture = true;
                continue;
            }
            if (capture && text.contains(end)) {
                break;
            }
            if (capture) {
                appendWithNewline(builder, text);
            }
        }
        return builder.toString();
    }

    private static List<SketchEntry> extractSketches(XWPFDocument document) {
        List<SketchEntry> sketches = new ArrayList<>();
        if (document == null) {
            return sketches;
        }
        XWPFTable table = findTableAfterTitle(document, "15. Эскизы (планы):");
        if (table == null) {
            return sketches;
        }
        for (XWPFTableRow row : table.getRows()) {
            if (row == null) {
                continue;
            }
            for (XWPFTableCell cell : row.getTableCells()) {
                if (cell == null) {
                    continue;
                }
                String caption = normalizeSpace(cell.getText());
                caption = caption.replaceFirst("(?i)^15\\.\\s*Эскизы\\s*\\(планы\\):?\\s*", "").trim();
                List<XWPFPictureData> pictures = extractPicturesFromCell(cell);
                if (pictures.isEmpty()) {
                    continue;
                }
                for (XWPFPictureData pictureData : pictures) {
                    if (pictureData == null || pictureData.getData() == null) {
                        continue;
                    }
                    sketches.add(new SketchEntry(pictureData.getData(), pictureData.getPictureType(), caption));
                }
            }
        }
        return sketches;
    }

    private static List<XWPFPictureData> extractPicturesFromCell(XWPFTableCell cell) {
        List<XWPFPictureData> pictures = new ArrayList<>();
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (org.apache.poi.xwpf.usermodel.XWPFRun run : paragraph.getRuns()) {
                for (XWPFPicture picture : run.getEmbeddedPictures()) {
                    XWPFPictureData data = picture.getPictureData();
                    if (data != null) {
                        pictures.add(data);
                    }
                }
            }
        }
        return pictures;
    }

    private static String extractTableTextAfterTitle(XWPFDocument document, String title) {
        XWPFTable table = findTableAfterTitle(document, title);
        if (table == null) {
            return "";
        }
        return tableToText(table);
    }

    private static String extractLineContaining(List<String> lines, String label) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String lowerLabel = label.toLowerCase(Locale.ROOT);
        for (String line : lines) {
            String normalized = normalizeSpace(line);
            if (normalized.toLowerCase(Locale.ROOT).contains(lowerLabel)) {
                return normalized;
            }
        }
        return "";
    }

    private static XWPFTable findTableAfterTitle(XWPFDocument document, String title) {
        if (document == null) {
            return null;
        }
        String lowerTitle = normalizeSpace(title).toLowerCase(Locale.ROOT);
        List<IBodyElement> elements = document.getBodyElements();
        for (int i = 0; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText()).toLowerCase(Locale.ROOT);
                if (!text.contains(lowerTitle)) {
                    continue;
                }
                for (int j = i + 1; j < elements.size(); j++) {
                    IBodyElement next = elements.get(j);
                    if (next instanceof XWPFTable table) {
                        return table;
                    }
                }
            } else if (element instanceof XWPFTable table) {
                if (tableContainsTitle(table, lowerTitle)) {
                    return table;
                }
            }
        }
        return null;
    }

    private static boolean tableContainsTitle(XWPFTable table, String lowerTitle) {
        if (table == null) {
            return false;
        }
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                String text = normalizeSpace(cell.getText()).toLowerCase(Locale.ROOT);
                if (text.contains(lowerTitle)) {
                    return true;
                }
            }
        }
        return false;
    }

    private static List<String> extractColumnValues(XWPFTable table, int columnIndex, int skipRows) {
        List<String> values = new ArrayList<>();
        if (table == null) {
            return values;
        }
        List<XWPFTableRow> rows = table.getRows();
        for (int rowIndex = skipRows; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = rows.get(rowIndex);
            if (row == null) {
                continue;
            }
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells == null || cells.size() <= columnIndex) {
                continue;
            }
            String value = normalizeSpace(cells.get(columnIndex).getText());
            if (!value.isBlank()) {
                values.add(value);
            }
        }
        return values;
    }

    private static List<InstrumentData> extractInstrumentRows(XWPFTable table, int nameCol, int serialCol) {
        List<InstrumentData> instruments = new ArrayList<>();
        if (table == null) {
            return instruments;
        }
        List<XWPFTableRow> rows = table.getRows();
        for (int rowIndex = 1; rowIndex < rows.size(); rowIndex++) {
            XWPFTableRow row = rows.get(rowIndex);
            if (row == null) {
                continue;
            }
            List<XWPFTableCell> cells = row.getTableCells();
            if (cells == null || cells.size() <= Math.max(nameCol, serialCol)) {
                continue;
            }
            String name = normalizeSpace(cells.get(nameCol).getText());
            String serial = normalizeSpace(cells.get(serialCol).getText());
            if (name.isBlank() && serial.isBlank()) {
                continue;
            }
            instruments.add(new InstrumentData(name, serial));
        }
        return instruments;
    }

    private static String tableToText(XWPFTable table) {
        if (table == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (XWPFTableRow row : table.getRows()) {
            List<String> cells = new ArrayList<>();
            for (XWPFTableCell cell : row.getTableCells()) {
                cells.add(normalizeSpace(cell.getText()));
            }
            String rowText = String.join(" | ", cells).trim();
            if (!rowText.isBlank()) {
                appendWithNewline(builder, rowText);
            }
        }
        return builder.toString();
    }

    private static void appendWithNewline(StringBuilder builder, String value) {
        if (builder.length() > 0) {
            builder.append('\n');
        }
        builder.append(value);
    }

    private static int writeConstructiveSolutionsTable(Sheet sheet, int rowIndex, String tableText) {
        if (tableText == null || tableText.isBlank()) {
            return rowIndex;
        }
        rowIndex = writeConstructiveSolutionsHeader(sheet, rowIndex);
        List<List<String>> rows = parseTableRows(tableText);
        for (List<String> row : rows) {
            if (isConstructiveSolutionsHeader(row)) {
                continue;
            }
            rowIndex = writeConstructiveSolutionsRow(sheet, rowIndex, row);
        }
        return rowIndex;
    }

    private static int writeRoomParametersTable(Sheet sheet, int rowIndex, String tableText) {
        if (tableText == null || tableText.isBlank()) {
            return rowIndex;
        }
        List<List<String>> rows = parseTableRows(tableText);
        for (List<String> row : rows) {
            rowIndex = writeRoomParametersRow(sheet, rowIndex, row);
        }
        return rowIndex;
    }

    private static int writeConstructiveSolutionsHeader(Sheet sheet, int rowIndex) {
        List<String> header = List.of("Тип конструкции", "Между помещениями", "Состав");
        return writeMergedRowWithColumns(sheet, rowIndex, header, new int[][]{{0, 6}, {7, 13}, {14, 31}}, true);
    }

    private static int writeConstructiveSolutionsRow(Sheet sheet, int rowIndex, List<String> columns) {
        List<String> values = normalizeToSize(columns, 3);
        return writeMergedRowWithColumns(sheet, rowIndex, values, new int[][]{{0, 6}, {7, 13}, {14, 31}}, true);
    }

    private static int writeRoomParametersRow(Sheet sheet, int rowIndex, List<String> columns) {
        List<String> values = normalizeToSize(columns, 4);
        return writeMergedRowWithColumns(sheet, rowIndex, values, new int[][]{{0, 6}, {7, 10}, {11, 14}, {15, 19}}, true);
    }

    private static int writeMergedRowWithColumns(Sheet sheet,
                                                 int rowIndex,
                                                 List<String> values,
                                                 int[][] ranges,
                                                 boolean addBorders) {
        if (sheet == null) {
            return rowIndex;
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        int totalFirstCol = ranges[0][0];
        int totalLastCol = ranges[ranges.length - 1][1];
        for (int index = 0; index < ranges.length; index++) {
            int startCol = ranges[index][0];
            int endCol = ranges[index][1];
            ensureMergedRegion(sheet, rowIndex, startCol, endCol);
            String value = index < values.size() ? safe(values.get(index)) : "";
            for (int col = startCol; col <= endCol; col++) {
                Cell cell = row.getCell(col);
                if (cell == null) {
                    cell = row.createCell(col);
                }
                if (col == startCol) {
                    cell.setCellValue(value);
                }
            }
            if (addBorders && value != null && !value.isBlank()) {
                applyBorderToMergedRegion(sheet, rowIndex, startCol, endCol);
            }
        }
        applyWrapStyleToRange(sheet, rowIndex, totalFirstCol, totalLastCol);
        String heightText = String.join(" ", values);
        adjustRowHeightForMergedText(sheet, rowIndex, totalFirstCol, totalLastCol, heightText);
        return rowIndex + 1;
    }

    private static void applyBorderToMergedRegion(Sheet sheet, int rowIndex, int startCol, int endCol) {
        if (sheet == null) {
            return;
        }
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, startCol, endCol);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private static List<List<String>> parseTableRows(String tableText) {
        List<List<String>> rows = new ArrayList<>();
        if (tableText == null || tableText.isBlank()) {
            return rows;
        }
        String[] lines = tableText.split("\\R");
        for (String line : lines) {
            String trimmed = line.trim();
            if (trimmed.isEmpty()) {
                continue;
            }
            String[] parts = trimmed.split("\\s*\\|\\s*");
            List<String> row = new ArrayList<>();
            for (String part : parts) {
                row.add(normalizeSpace(part));
            }
            rows.add(row);
        }
        return rows;
    }

    private static List<String> normalizeToSize(List<String> values, int targetSize) {
        List<String> normalized = new ArrayList<>();
        if (values != null) {
            normalized.addAll(values);
        }
        if (normalized.size() > targetSize) {
            List<String> trimmed = new ArrayList<>(normalized.subList(0, targetSize - 1));
            String tail = String.join(" ", normalized.subList(targetSize - 1, normalized.size()));
            trimmed.add(tail);
            return trimmed;
        }
        while (normalized.size() < targetSize) {
            normalized.add("");
        }
        return normalized;
    }

    private static boolean isConstructiveSolutionsHeader(List<String> row) {
        if (row == null || row.isEmpty()) {
            return false;
        }
        String joined = String.join(" ", row).toLowerCase(Locale.ROOT);
        return joined.contains("тип конструкции") || joined.contains("между помещениями") || joined.contains("состав");
    }

    private static class LnwAcSourceData {
        private Sheet sourceSheet;
        private int startRow = -1;
        private int endRow = -1;
        private String firstRoomWord = "";
        private String secondRoomWord = "";
    }

    private record SoundInsulationProtocolData(String registrationNumber,
                                               String customerNameAndContacts,
                                               String measurementDates,
                                               String measurementPerformer,
                                               String representative,
                                               String controlPerson,
                                               String controlDate,
                                               String protocolNumber,
                                               String contractText,
                                               String customerLegalAddress,
                                               String objectName,
                                               String objectAddress,
                                               String measurementMethods,
                                               List<InstrumentData> instruments,
                                               List<String> roomNames,
                                               String objectDetails,
                                               String constructiveSolutionsTable,
                                               String roomParametersTable,
                                               String areaBetweenRooms,
                                               List<SketchEntry> sketches) {
        private SoundInsulationProtocolData() {
            this("", "", "", "", "", "", "", "", "", "", "", "", "", new ArrayList<>(), new ArrayList<>(),
                    "", "", "", "", new ArrayList<>());
        }
    }

    private record InstrumentData(String name, String serialNumber) {
    }

    private record SketchEntry(byte[] imageData, int pictureType, String caption) {
    }
}
