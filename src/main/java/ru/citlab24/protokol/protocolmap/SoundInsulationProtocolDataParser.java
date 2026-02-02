package ru.citlab24.protokol.protocolmap;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

final class SoundInsulationProtocolDataParser {
    private static final String MEASUREMENT_DATES_LABEL = "Дата проведения измерений:";
    private static final String OBJECT_NAME_LABEL =
            "4. Наименование предприятия, организации, объекта, где производились измерения:";
    private static final String OBJECT_ADDRESS_LABEL = "5. Адрес предприятия (объекта):";
    private static final String CUSTOMER_LABEL = "1. Наименование и контактные данные заявителя (заказчика):";
    private static final String PROTOCOL_NUMBER_LABEL = "Протокол испытаний №";
    private static final String REGISTRATION_LABEL = "9. Регистрационный номер карты:";
    private static final String MEASUREMENT_BASIS_LABEL = "Основание для измерений:";
    private static final String INSTRUMENTS_HEADER = "Наименование, тип средства измерения";
    private static final String EQUIPMENT_SECTION_TITLE = "11. Сведения об испытательном оборудовании:";
    private static final String METHODS_TITLE =
            "12. Сведения о нормативных документах (НД), регламентирующих значения показателей и НД " +
                    "на методы (методики) измерений:";

    private SoundInsulationProtocolDataParser() {
    }

    static ProtocolData parse(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return ProtocolData.empty();
        }
        if (!protocolFile.getName().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            return ProtocolData.empty();
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document, true);
            List<String> paragraphLines = extractLines(document, false);
            String measurementDatesRaw = extractValueAfterLabel(lines, MEASUREMENT_DATES_LABEL);
            List<String> measurementDates = splitDates(measurementDatesRaw);
            String objectName = extractValueBetweenLabels(lines, OBJECT_NAME_LABEL, OBJECT_ADDRESS_LABEL);
            String objectAddress = extractValueAfterLabel(lines, OBJECT_ADDRESS_LABEL);
            String customerName = extractValueAfterLabel(lines, CUSTOMER_LABEL);
            customerName = trimToComma(customerName);
            String protocolNumber = extractValueAfterLabel(lines, PROTOCOL_NUMBER_LABEL);
            String registrationNumber = extractValueAfterLabel(lines, REGISTRATION_LABEL);
            String applicationNumber = extractApplicationNumber(lines);
            String protocolDate = extractProtocolDate(paragraphLines);
            List<InstrumentEntry> instruments = extractInstruments(document);
            instruments.addAll(extractEquipmentInstruments(document));
            String measurementMethods = extractMeasurementMethods(document);
            return new ProtocolData(protocolNumber, protocolDate, customerName, registrationNumber, applicationNumber,
                    objectName, objectAddress, measurementDates, instruments, measurementMethods);
        } catch (Exception ex) {
            return ProtocolData.empty();
        }
    }

    private static List<String> extractLines(XWPFDocument document, boolean includeTables) {
        List<String> lines = new ArrayList<>();
        if (document == null) {
            return lines;
        }
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                addParagraphLines(lines, paragraph.getText());
            } else if (includeTables && element instanceof XWPFTable table) {
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
            String value = builder.toString().trim();
            if (!value.isBlank()) {
                lines.add(value);
            }
        }
    }

    private static String extractValueAfterLabel(List<String> lines, String label) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String normalizedLabel = normalizeSpace(label).toLowerCase(Locale.ROOT);
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            String normalizedLine = normalizeSpace(line);
            String lowerLine = normalizedLine.toLowerCase(Locale.ROOT);
            if (!lowerLine.contains(normalizedLabel)) {
                continue;
            }
            int index = lowerLine.indexOf(normalizedLabel);
            String tail = normalizedLine.substring(index + normalizedLabel.length()).trim();
            if (!tail.isBlank()) {
                return stripLeadingSeparators(tail);
            }
            for (int j = i + 1; j < lines.size(); j++) {
                String next = normalizeSpace(lines.get(j));
                if (!next.isBlank()) {
                    return stripLeadingSeparators(next);
                }
            }
        }
        return "";
    }

    private static String extractValueBetweenLabels(List<String> lines, String startLabel, String endLabel) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String normalizedStart = normalizeSpace(startLabel).toLowerCase(Locale.ROOT);
        String normalizedEnd = normalizeSpace(endLabel).toLowerCase(Locale.ROOT);
        StringBuilder result = new StringBuilder();
        boolean collecting = false;
        for (String line : lines) {
            String normalizedLine = normalizeSpace(line);
            String lowerLine = normalizedLine.toLowerCase(Locale.ROOT);
            if (!collecting) {
                int startIndex = lowerLine.indexOf(normalizedStart);
                if (startIndex < 0) {
                    continue;
                }
                String tail = normalizedLine.substring(startIndex + normalizedStart.length()).trim();
                collecting = true;
                if (!tail.isBlank()) {
                    if (appendUntilEndLabel(result, tail, normalizedEnd)) {
                        break;
                    }
                }
                continue;
            }
            if (appendUntilEndLabel(result, normalizedLine, normalizedEnd)) {
                break;
            }
        }
        return stripLeadingSeparators(result.toString().trim());
    }

    private static boolean appendUntilEndLabel(StringBuilder builder, String value, String normalizedEndLabel) {
        String lowerValue = value.toLowerCase(Locale.ROOT);
        int endIndex = lowerValue.indexOf(normalizedEndLabel);
        if (endIndex >= 0) {
            String chunk = value.substring(0, endIndex).trim();
            appendWithSpace(builder, chunk);
            return true;
        }
        appendWithSpace(builder, value);
        return false;
    }

    private static void appendWithSpace(StringBuilder builder, String value) {
        if (value.isBlank()) {
            return;
        }
        if (builder.length() > 0) {
            builder.append(' ');
        }
        builder.append(value);
    }

    private static String extractApplicationNumber(List<String> lines) {
        String value = extractValueAfterLabel(lines, MEASUREMENT_BASIS_LABEL);
        if (value.isBlank()) {
            return "";
        }
        String lower = value.toLowerCase(Locale.ROOT);
        int index = lower.indexOf("заявка");
        if (index >= 0) {
            String tail = value.substring(index + "заявка".length()).trim();
            return trimLeadingPunctuation(tail);
        }
        return trimLeadingPunctuation(value);
    }

    private static String extractProtocolDate(List<String> lines) {
        if (lines == null) {
            return "";
        }
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line.contains("/М.О. Тарновский/") || line.contains("М.О. Тарновский")) {
                for (int j = i + 1; j < lines.size(); j++) {
                    String next = normalizeSpace(lines.get(j));
                    if (!next.isBlank()) {
                        return next;
                    }
                }
            }
        }
        return "";
    }

    private static List<InstrumentEntry> extractInstruments(XWPFDocument document) {
        List<InstrumentEntry> instruments = new ArrayList<>();
        if (document == null) {
            return instruments;
        }
        for (XWPFTable table : document.getTables()) {
            InstrumentHeader header = findInstrumentHeader(table);
            if (header == null) {
                continue;
            }
            instruments.addAll(readInstrumentRows(table, header));
            if (!instruments.isEmpty()) {
                return instruments;
            }
        }
        return instruments;
    }

    private static List<InstrumentEntry> extractEquipmentInstruments(XWPFDocument document) {
        List<InstrumentEntry> instruments = new ArrayList<>();
        if (document == null) {
            return instruments;
        }
        XWPFTable table = findTableAfterTitle(document, EQUIPMENT_SECTION_TITLE);
        if (table == null) {
            return instruments;
        }
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            List<XWPFTableCell> cells = row.getTableCells();
            int nameColumn = findColumnByHeader(cells, "наименование");
            if (nameColumn < 0) {
                continue;
            }
            int serialColumn = findColumnByHeader(cells, "зав");
            if (serialColumn < 0 && nameColumn + 1 < cells.size()) {
                serialColumn = nameColumn + 1;
            }
            for (int dataRowIndex = rowIndex + 1; dataRowIndex < table.getNumberOfRows(); dataRowIndex++) {
                XWPFTableRow dataRow = table.getRow(dataRowIndex);
                if (dataRow == null) {
                    continue;
                }
                String name = getCellText(dataRow, nameColumn);
                String serial = getCellText(dataRow, serialColumn);
                if (name.isBlank() && serial.isBlank()) {
                    if (rowIsEmpty(dataRow)) {
                        break;
                    }
                    continue;
                }
                instruments.add(new InstrumentEntry(name, serial));
            }
            if (!instruments.isEmpty()) {
                return instruments;
            }
        }
        return instruments;
    }

    private static int findColumnByHeader(List<XWPFTableCell> cells, String header) {
        if (cells == null) {
            return -1;
        }
        String normalizedHeader = header.toLowerCase(Locale.ROOT);
        for (int colIndex = 0; colIndex < cells.size(); colIndex++) {
            String text = normalizeSpace(cells.get(colIndex).getText()).toLowerCase(Locale.ROOT);
            if (text.contains(normalizedHeader)) {
                return colIndex;
            }
        }
        return -1;
    }

    private static InstrumentHeader findInstrumentHeader(XWPFTable table) {
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            List<XWPFTableCell> cells = row.getTableCells();
            for (int colIndex = 0; colIndex < cells.size(); colIndex++) {
                String text = normalizeSpace(cells.get(colIndex).getText()).toLowerCase(Locale.ROOT);
                if (!text.contains(INSTRUMENTS_HEADER.toLowerCase(Locale.ROOT))) {
                    continue;
                }
                int serialCol = findSerialColumn(cells);
                if (serialCol < 0 && colIndex + 1 < cells.size()) {
                    serialCol = colIndex + 1;
                }
                return new InstrumentHeader(rowIndex, colIndex, serialCol);
            }
        }
        return null;
    }

    private static int findSerialColumn(List<XWPFTableCell> cells) {
        for (int colIndex = 0; colIndex < cells.size(); colIndex++) {
            String text = normalizeSpace(cells.get(colIndex).getText()).toLowerCase(Locale.ROOT);
            if (text.contains("зав")) {
                return colIndex;
            }
        }
        return -1;
    }

    private static List<InstrumentEntry> readInstrumentRows(XWPFTable table, InstrumentHeader header) {
        List<InstrumentEntry> instruments = new ArrayList<>();
        for (int rowIndex = header.rowIndex + 1; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row == null) {
                continue;
            }
            String name = getCellText(row, header.nameColumn);
            String serial = getCellText(row, header.serialColumn);
            if (name.isBlank() && serial.isBlank()) {
                if (rowIsEmpty(row)) {
                    break;
                }
                continue;
            }
            instruments.add(new InstrumentEntry(name, serial));
        }
        return instruments;
    }

    private static boolean rowIsEmpty(XWPFTableRow row) {
        for (XWPFTableCell cell : row.getTableCells()) {
            if (!normalizeSpace(cell.getText()).isBlank()) {
                return false;
            }
        }
        return true;
    }

    private static String getCellText(XWPFTableRow row, int columnIndex) {
        if (columnIndex < 0) {
            return "";
        }
        List<XWPFTableCell> cells = row.getTableCells();
        if (columnIndex >= cells.size()) {
            return "";
        }
        return normalizeSpace(cells.get(columnIndex).getText());
    }

    private static String extractMeasurementMethods(XWPFDocument document) {
        XWPFTable table = findTableAfterTitle(document, METHODS_TITLE);
        if (table == null) {
            return "";
        }
        List<String> values = extractColumnValues(table, 2, 1);
        return String.join("; ", values);
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

    private static List<String> splitDates(String rawDates) {
        List<String> dates = new ArrayList<>();
        if (rawDates == null || rawDates.isBlank()) {
            return dates;
        }
        for (String part : rawDates.split(",")) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) {
                dates.add(trimmed);
            }
        }
        return dates;
    }

    private static String trimToComma(String value) {
        if (value == null) {
            return "";
        }
        int commaIndex = value.indexOf(',');
        if (commaIndex >= 0) {
            return value.substring(0, commaIndex).trim();
        }
        return value.trim();
    }

    private static String trimLeadingPunctuation(String value) {
        if (value == null) {
            return "";
        }
        int index = 0;
        while (index < value.length() && !Character.isLetterOrDigit(value.charAt(index))) {
            index++;
        }
        return value.substring(index).trim();
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

    private record InstrumentHeader(int rowIndex, int nameColumn, int serialColumn) {
    }

    record InstrumentEntry(String name, String serialNumber) {
    }

    record ProtocolData(String protocolNumber,
                        String protocolDate,
                        String customerName,
                        String registrationNumber,
                        String applicationNumber,
                        String objectName,
                        String objectAddress,
                        List<String> measurementDates,
                        List<InstrumentEntry> instruments,
                        String measurementMethods) {
        private static ProtocolData empty() {
            return new ProtocolData("", "", "", "", "", "", "", new ArrayList<>(), new ArrayList<>(), "");
        }
    }
}
