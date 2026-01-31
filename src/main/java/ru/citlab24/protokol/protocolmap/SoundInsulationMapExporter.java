package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
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
        if (protocolFile == null || !protocolFile.exists()) {
            return targetFile;
        }
        try {
            SoundInsulationProtocolData data = extractProtocolData(protocolFile);
            applyProtocolData(targetFile, data);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return targetFile;
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
            String measurementPerformer = resolveMeasurementPerformer(lines);
            String controlPerson = resolveControlPerson(measurementPerformer);
            return new SoundInsulationProtocolData(registrationNumber, customer, measurementDates,
                    measurementPerformer, representative, controlPerson, controlDate, protocolNumber, contractText,
                    legalAddress, objectName, objectAddress);
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

    private static String resolveMeasurementPerformer(List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return "заведующий лабораторией Тарновский М.О.";
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
            return "заведующий лабораторией Тарновский М.О.";
        }
        return "заведующий лабораторией Тарновский М.О.";
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
                                               String objectAddress) {
        private SoundInsulationProtocolData() {
            this("", "", "", "", "", "", "", "", "", "", "", "");
        }
    }
}
