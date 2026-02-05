package ru.citlab24.protokol.protocolmap;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;

final class SoundInsulationRequestFormExporter {
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int PLAN_TABLE_FONT_SIZE = 10;
    private static final double REQUEST_TABLE_WIDTH_SCALE = 0.8;
    private static final String PLAN_TABLE_HEADER = "Наименование объекта измерений";
    private static final String CUSTOMER_SECTION_HEADER =
            "Данные, предоставленные Заказчиком, за которые он несет ответственность";
    private static final String PLAN_APPENDIX_TITLE = "ПЛАН ИЗМЕРЕНИЙ";
    private static final String PLAN_APPENDIX_HEADER = "Приложение – План измерений";
    private static final String CUSTOMER_APPENDIX_HEADER = "Приложение – Данные, предоставлены Заказчиком";
    private static final String CUSTOMER_DATA_START =
            "Объект испытаний – внутренние ограждающие конструкции помещений, " +
                    "их монтаж осуществлен заказчиком согласно требованиям технической документации.";
    private static final String CUSTOMER_DATA_END =
            "14. Сведения о дополнении, отклонении или исключении из методов: -";
    private static final String ROOM_PARAMS_START = "Параметры помещений и испытываемой поверхности:";
    private static final String ROOM_PARAMS_END = "17. Результаты измерений";
    private static final String ROOM_PARAMS_ALT_START = "16. Параметры помещений и испытываемой поверхности:";
    private static final String AREA_BETWEEN_ROOMS_MARKER =
            "Площадь испытываемой поверхности между помещениями";
    private static final String NORMATIVE_REQUIREMENTS_MARKER = "нормативные требования";
    private static final Pattern BRACKETS_PATTERN = Pattern.compile("\\([^)]*\\)");
    private static final String APPLICATION_BASIS_LABEL = "6. Основание для измерений";
    private static final String APPLICATION_BASIS_ALT_LABEL = "Основание для измерений";

    private SoundInsulationRequestFormExporter() {
    }

    static void generate(File protocolFile, File mapFile, String workDeadline, String customerInn) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveRequestFormFile(mapFile);
        String applicationNumber = resolveApplicationNumber(protocolFile, mapFile);
        File planFile = SoundInsulationMeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
        List<List<String>> planRows = extractPlanRows(planFile);
        List<String> customerLines = extractCustomerDataLines(protocolFile);
        List<List<String>> roomParamsRows = extractRoomParamsTable(protocolFile);
        List<String> roomParamsLines = roomParamsRows.isEmpty() ? extractRoomParamsLines(protocolFile) : List.of();
        String areaBetweenRoomsLine = extractAreaBetweenRoomsLine(protocolFile);
        List<String> measurementRequirementLines = extractMeasurementRequirementLines(protocolFile);

        if (planRows.isEmpty()) {
            List<String> empty = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                empty.add("");
            }
            planRows.add(empty);
        }
        applyPlanDeadline(planRows, workDeadline);

        try (XWPFDocument document = new XWPFDocument()) {
            RequestFormExporter.applyStandardHeader(document);

            XWPFParagraph appendixHeader = document.createParagraph();
            appendixHeader.setAlignment(ParagraphAlignment.RIGHT);
            setParagraphSpacing(appendixHeader);
            XWPFRun appendixHeaderRun = appendixHeader.createRun();
            appendixHeaderRun.setFontFamily(FONT_NAME);
            appendixHeaderRun.setFontSize(FONT_SIZE);
            setRunTextWithBreaks(appendixHeaderRun,
                    "Приложения к заявке № " + applicationNumber + "\n" + PLAN_APPENDIX_HEADER);

            XWPFParagraph appendixTitle = document.createParagraph();
            appendixTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(appendixTitle);
            XWPFRun appendixTitleRun = appendixTitle.createRun();
            appendixTitleRun.setText(PLAN_APPENDIX_TITLE);
            appendixTitleRun.setFontFamily(FONT_NAME);
            appendixTitleRun.setFontSize(FONT_SIZE);
            appendixTitleRun.setBold(true);

            XWPFTable planTable = document.createTable(planRows.size(), 5);
            configureTableLayout(planTable, new int[]{2500, 2500, 2100, 2560, 2900});
            for (int rowIndex = 0; rowIndex < planRows.size(); rowIndex++) {
                List<String> row = planRows.get(rowIndex);
                boolean isHeader = rowIndex == 0;
                ParagraphAlignment alignment = isHeader ? ParagraphAlignment.CENTER : ParagraphAlignment.LEFT;
                for (int colIndex = 0; colIndex < 5; colIndex++) {
                    String value = colIndex < row.size() ? row.get(colIndex) : "";
                    setTableCellText(planTable.getRow(rowIndex).getCell(colIndex), value,
                            PLAN_TABLE_FONT_SIZE, isHeader, alignment);
                }
            }

            XWPFParagraph spacerAfterPlan = document.createParagraph();
            setParagraphSpacing(spacerAfterPlan);

            addParagraphWithLineBreaks(document,
                    "Представитель заказчика _______________________________________________\n" +
                            "                                                 (Должность, ФИО, контактные данные)  ");

            addPageBreak(document);

            XWPFParagraph customerAppendixHeader = document.createParagraph();
            customerAppendixHeader.setAlignment(ParagraphAlignment.RIGHT);
            setParagraphSpacing(customerAppendixHeader);
            XWPFRun customerAppendixHeaderRun = customerAppendixHeader.createRun();
            customerAppendixHeaderRun.setFontFamily(FONT_NAME);
            customerAppendixHeaderRun.setFontSize(FONT_SIZE);
            setRunTextWithBreaks(customerAppendixHeaderRun,
                    "Приложения к заявке № " + applicationNumber + "\n" + CUSTOMER_APPENDIX_HEADER);

            XWPFParagraph customerAppendixTitle = document.createParagraph();
            customerAppendixTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(customerAppendixTitle);
            XWPFRun customerAppendixTitleRun = customerAppendixTitle.createRun();
            customerAppendixTitleRun.setText(CUSTOMER_SECTION_HEADER);
            customerAppendixTitleRun.setFontFamily(FONT_NAME);
            customerAppendixTitleRun.setFontSize(FONT_SIZE);
            customerAppendixTitleRun.setBold(true);

            for (String line : customerLines) {
                addParagraphWithLineBreaks(document, line);
            }

            if (!roomParamsRows.isEmpty()) {
                addParagraphWithLineBreaks(document, ROOM_PARAMS_START);
                XWPFTable roomParamsTable = document.createTable(roomParamsRows.size(), 4);
                configureTableLayout(roomParamsTable, new int[]{3600, 2600, 2000, 2200});
                for (int rowIndex = 0; rowIndex < roomParamsRows.size(); rowIndex++) {
                    List<String> row = roomParamsRows.get(rowIndex);
                    boolean isHeader = rowIndex == 0;
                    ParagraphAlignment alignment = isHeader ? ParagraphAlignment.CENTER : ParagraphAlignment.LEFT;
                    for (int colIndex = 0; colIndex < 4; colIndex++) {
                        String value = colIndex < row.size() ? row.get(colIndex) : "";
                        setTableCellText(roomParamsTable.getRow(rowIndex).getCell(colIndex), value,
                                PLAN_TABLE_FONT_SIZE, isHeader, alignment);
                    }
                }
            } else {
                for (String line : roomParamsLines) {
                    addParagraphWithLineBreaks(document, line);
                }
            }

            if (!areaBetweenRoomsLine.isBlank()) {
                addParagraphWithLineBreaks(document, areaBetweenRoomsLine);
            }

            for (String requirementLine : measurementRequirementLines) {
                addParagraphWithLineBreaks(document, requirementLine);
            }

            XWPFParagraph spacerBeforeSignature = document.createParagraph();
            setParagraphSpacing(spacerBeforeSignature);

            addParagraphWithLineBreaks(document,
                    "Представитель заказчика _______________________________________________\n" +
                            "(Должность, ФИО, контактные данные)\n");

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveRequestFormFile(File mapFile) {
        return RequestFormExporter.resolveRequestFormFile(mapFile);
    }

    private static String resolveApplicationNumber(File protocolFile, File mapFile) {
        String fromProtocol = extractApplicationNumberFromProtocol(protocolFile);
        if (!fromProtocol.isBlank()) {
            return fromProtocol;
        }
        return RequestFormExporter.resolveApplicationNumberFromMap(mapFile);
    }

    private static String extractApplicationNumberFromProtocol(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return "";
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document, true);
            return extractApplicationNumberFromLines(lines);
        } catch (Exception ignored) {
            return "";
        }
    }

    private static String extractApplicationNumberFromLines(List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String basisLabel = APPLICATION_BASIS_LABEL.toLowerCase(Locale.ROOT);
        String basisAltLabel = APPLICATION_BASIS_ALT_LABEL.toLowerCase(Locale.ROOT);
        for (String line : lines) {
            String normalized = normalizeSpace(line);
            String lower = normalized.toLowerCase(Locale.ROOT);
            if (!lower.contains(basisLabel) && !lower.contains(basisAltLabel)) {
                continue;
            }
            int applicationIndex = lower.indexOf("заявка");
            if (applicationIndex < 0) {
                continue;
            }
            String tail = normalized.substring(applicationIndex + "заявка".length()).trim();
            return trimLeadingPunctuation(tail);
        }
        return "";
    }

    private static List<List<String>> extractPlanRows(File planFile) {
        List<List<String>> rows = new ArrayList<>();
        if (planFile == null || !planFile.exists()) {
            return rows;
        }
        try (InputStream inputStream = new FileInputStream(planFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            for (XWPFTable table : document.getTables()) {
                if (!isPlanTable(table)) {
                    continue;
                }
                for (XWPFTableRow row : table.getRows()) {
                    List<String> values = new ArrayList<>();
                    values.add(getCellText(row, 0));
                    values.add(getCellText(row, 1));
                    values.add(getCellText(row, 3));
                    values.add(getCellText(row, 4));
                    values.add(getCellText(row, 5));
                    rows.add(values);
                }
                break;
            }
        } catch (Exception ignored) {
            // пропускаем чтение плана измерений при ошибке
        }
        return rows;
    }

    private static void applyPlanDeadline(List<List<String>> rows, String workDeadline) {
        if (rows == null || rows.size() < 2) {
            return;
        }
        if (workDeadline == null || workDeadline.isBlank()) {
            return;
        }
        List<String> row = rows.get(1);
        while (row.size() < 5) {
            row.add("");
        }
        row.set(4, workDeadline.trim());
    }

    private static boolean isPlanTable(XWPFTable table) {
        if (table == null || table.getNumberOfRows() == 0) {
            return false;
        }
        XWPFTableRow row = table.getRow(0);
        if (row == null || row.getTableCells().size() < 6) {
            return false;
        }
        String header = normalizeSpace(getCellText(row, 0));
        return header.toLowerCase(Locale.ROOT).contains(PLAN_TABLE_HEADER.toLowerCase(Locale.ROOT));
    }

    private static String getCellText(XWPFTableRow row, int index) {
        if (row == null || row.getTableCells().size() <= index) {
            return "";
        }
        XWPFTableCell cell = row.getCell(index);
        return cell == null ? "" : normalizeSpace(cell.getText());
    }

    private static List<String> extractCustomerDataLines(File protocolFile) {
        List<String> lines = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return lines;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            lines.addAll(extractSection(allLines, CUSTOMER_DATA_START, CUSTOMER_DATA_END));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return lines;
    }

    private static List<String> extractRoomParamsLines(File protocolFile) {
        List<String> lines = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return lines;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            lines.addAll(extractSection(allLines, ROOM_PARAMS_START, ROOM_PARAMS_END));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return lines;
    }

    private static String extractAreaBetweenRoomsLine(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return "";
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            for (String line : allLines) {
                String normalized = normalizeSpace(line);
                if (normalized.toLowerCase(Locale.ROOT)
                        .contains(AREA_BETWEEN_ROOMS_MARKER.toLowerCase(Locale.ROOT))) {
                    return normalized;
                }
            }
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return "";
    }

    private static List<List<String>> extractRoomParamsTable(File protocolFile) {
        List<List<String>> rows = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return rows;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            XWPFTable table = findTableAfterTitle(document, ROOM_PARAMS_START);
            if (table == null) {
                table = findTableAfterTitle(document, ROOM_PARAMS_ALT_START);
            }
            if (table == null) {
                return rows;
            }
            for (XWPFTableRow row : table.getRows()) {
                List<String> values = new ArrayList<>();
                for (int colIndex = 0; colIndex < 4; colIndex++) {
                    values.add(getCellText(row, colIndex));
                }
                rows.add(values);
            }
        } catch (Exception ignored) {
            // пропускаем извлечение таблицы при ошибке
        }
        return rows;
    }

    private static List<String> extractSection(List<String> lines, String startMarker, String endMarker) {
        List<String> result = new ArrayList<>();
        if (lines == null || lines.isEmpty()) {
            return result;
        }
        String lowerStart = startMarker.toLowerCase(Locale.ROOT);
        String lowerEnd = endMarker.toLowerCase(Locale.ROOT);
        boolean collecting = false;
        for (String line : lines) {
            String normalized = normalizeSpace(line);
            String lower = normalized.toLowerCase(Locale.ROOT);
            if (!collecting) {
                int startIndex = lower.indexOf(lowerStart);
                if (startIndex < 0) {
                    continue;
                }
                String value = normalized.substring(startIndex);
                collecting = true;
                if (appendUntilEnd(result, value, lowerEnd)) {
                    break;
                }
                continue;
            }
            if (appendUntilEnd(result, normalized, lowerEnd)) {
                break;
            }
        }
        return sanitizeCustomerLines(result);
    }

    private static boolean appendUntilEnd(List<String> result, String value, String lowerEndMarker) {
        if (value == null) {
            return false;
        }
        String lower = value.toLowerCase(Locale.ROOT);
        int endIndex = lower.indexOf(lowerEndMarker);
        if (endIndex >= 0) {
            String fragment = value.substring(0, endIndex).trim();
            if (!fragment.isBlank()) {
                result.add(fragment);
            }
            return true;
        }
        if (!value.isBlank()) {
            result.add(value);
        }
        return false;
    }

    private static List<String> sanitizeCustomerLines(List<String> lines) {
        List<String> sanitized = new ArrayList<>();
        for (String line : lines) {
            if (line == null || line.isBlank()) {
                continue;
            }
            String value = line;
            if (value.toLowerCase(Locale.ROOT).contains("13.1 конструктивные решения")) {
                value = value.replaceFirst("(?i)^13\\.1\\s+", "");
            }
            sanitized.add(value);
        }
        return sanitized;
    }

    private static List<String> extractLines(XWPFDocument document, boolean includeTables) {
        List<String> lines = new ArrayList<>();
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

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
    }

    private static void addParagraphWithLineBreaks(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
        setRunTextWithBreaks(run, text);
    }

    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph pageBreakParagraph = document.createParagraph();
        setParagraphSpacing(pageBreakParagraph);
        pageBreakParagraph.setPageBreak(true);
    }

    private static List<String> extractMeasurementRequirementLines(File protocolFile) {
        List<String> result = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return result;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document, true);
            result.addAll(extractLnRequirements(lines));
            result.addAll(extractRRequirements(lines));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return result;
    }

    private static List<String> extractLnRequirements(List<String> lines) {
        List<String> result = new ArrayList<>();
        int start = findLineIndexContaining(lines, "17.1");
        int end = findLineIndexContaining(lines, "17.2");
        if (start < 0) {
            return result;
        }
        if (end < 0) {
            end = lines.size();
        }
        for (int i = start; i < end; i++) {
            if (!containsIgnoreCase(lines.get(i), "Ln, дБ")) {
                continue;
            }
            String source = buildSearchWindow(lines, i + 1, end);
            String rooms = extractBetweenMarkers(source,
                    "Индекс приведенного уровня ударного шума (Lnw) по результатам измерений для перекрытия между помещениями",
                    NORMATIVE_REQUIREMENTS_MARKER);
            String requirements = extractTailFromMarker(source, NORMATIVE_REQUIREMENTS_MARKER, false);
            if (rooms.isBlank() || requirements.isBlank()) {
                continue;
            }
            result.add("Для перекрытия между помещениями " + rooms + " нормативные требования " + requirements);
        }
        return result;
    }

    private static List<String> extractRRequirements(List<String> lines) {
        List<String> result = new ArrayList<>();
        int start = findLineIndexContaining(lines, "17.2");
        int end = findLineIndexContaining(lines, "18.");
        if (start < 0) {
            return result;
        }
        if (end < 0) {
            end = lines.size();
        }
        for (int i = start; i < end; i++) {
            if (!isRwTableHeader(lines.get(i))) {
                continue;
            }
            String source = buildSearchWindow(lines, i + 1, findNextRwHeaderIndex(lines, i + 1, end));
            String partitionRooms = extractBetweenMarkers(source,
                    "по результатам измерений для перегородки между помещениями",
                    NORMATIVE_REQUIREMENTS_MARKER);
            partitionRooms = takeFirstWords(partitionRooms, 3);
            String requirements = extractTailFromMarker(source, NORMATIVE_REQUIREMENTS_MARKER, true);
            if (!partitionRooms.isBlank() && !requirements.isBlank()) {
                result.add("Для перегородки между помещениями "
                        + partitionRooms + ": нормативные требования " + requirements);
            }
        }
        return result;
    }

    private static String takeFirstWords(String value, int wordsCount) {
        String normalized = normalizeSpace(value);
        if (normalized.isBlank() || wordsCount <= 0) {
            return "";
        }
        String[] words = normalized.split("\\s+");
        if (words.length <= wordsCount) {
            return normalized;
        }
        return String.join(" ", Arrays.copyOfRange(words, 0, wordsCount));
    }

    private static boolean isRwTableHeader(String line) {
        return containsIgnoreCase(line, "R, дБ") || containsIgnoreCase(line, "Rw, дБ");
    }

    private static int findNextRwHeaderIndex(List<String> lines, int fromIndex, int endExclusive) {
        int max = Math.min(endExclusive, lines.size());
        for (int i = Math.max(0, fromIndex); i < max; i++) {
            if (isRwTableHeader(lines.get(i))) {
                return i;
            }
        }
        return max;
    }

    private static int findLineIndexContaining(List<String> lines, String marker) {
        if (lines == null || marker == null || marker.isBlank()) {
            return -1;
        }
        for (int i = 0; i < lines.size(); i++) {
            if (containsIgnoreCase(lines.get(i), marker)) {
                return i;
            }
        }
        return -1;
    }

    private static String buildSearchWindow(List<String> lines, int start, int end) {
        StringBuilder builder = new StringBuilder();
        int max = Math.min(end, start + 24);
        for (int i = start; i < max; i++) {
            String value = normalizeSpace(lines.get(i));
            if (value.matches("^\\d+\\.\\d+.*") || value.matches("^\\d+\\..*")) {
                break;
            }
            if (builder.length() > 0) {
                builder.append(' ');
            }
            builder.append(value);
        }
        return builder.toString();
    }

    private static String extractBetweenMarkers(String source, String startMarker, String endMarker) {
        if (source == null || source.isBlank()) {
            return "";
        }
        String lower = source.toLowerCase(Locale.ROOT);
        String lowerStart = startMarker.toLowerCase(Locale.ROOT);
        String lowerEnd = endMarker.toLowerCase(Locale.ROOT);
        int startIndex = lower.indexOf(lowerStart);
        if (startIndex < 0) {
            return "";
        }
        int fromIndex = startIndex + lowerStart.length();
        int endIndex = lower.indexOf(lowerEnd, fromIndex);
        if (endIndex < 0) {
            return "";
        }
        return normalizeSpace(source.substring(fromIndex, endIndex));
    }

    private static String extractTailFromMarker(String source, String marker, boolean removeBrackets) {
        if (source == null || source.isBlank()) {
            return "";
        }
        String lower = source.toLowerCase(Locale.ROOT);
        String lowerMarker = marker.toLowerCase(Locale.ROOT);
        int index = lower.indexOf(lowerMarker);
        if (index < 0) {
            return "";
        }
        String tail = source.substring(index + marker.length());
        if (removeBrackets) {
            tail = BRACKETS_PATTERN.matcher(tail).replaceAll(" ");
        }
        tail = normalizeSpace(tail);
        while (tail.startsWith(":")) {
            tail = normalizeSpace(tail.substring(1));
        }
        return tail;
    }

    private static boolean containsIgnoreCase(String value, String marker) {
        if (value == null || marker == null) {
            return false;
        }
        return normalizeSpace(value).toLowerCase(Locale.ROOT)
                .contains(normalizeSpace(marker).toLowerCase(Locale.ROOT));
    }

    private static void setRunTextWithBreaks(XWPFRun run, String text) {
        if (text == null) {
            return;
        }
        String[] parts = text.split("\\n", -1);
        int textPos = 0;
        for (int index = 0; index < parts.length; index++) {
            if (index > 0) {
                run.addBreak();
            }
            run.setText(parts[index], textPos++);
        }
    }

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(scaleWidth(12560)));

        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable.CENTER);

        CTTblLayoutType layout = pr.isSetTblLayout() ? pr.getTblLayout() : pr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);

        CTTblGrid grid = ct.getTblGrid();
        if (grid == null) {
            grid = ct.addNewTblGrid();
        } else {
            while (grid.sizeOfGridColArray() > 0) {
                grid.removeGridCol(0);
            }
        }
        int[] scaledWidths = scaleColumnWidths(columnWidths);
        for (int width : scaledWidths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }

        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < scaledWidths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, scaledWidths[colIndex]);
            }
        }
    }

    private static int[] scaleColumnWidths(int[] columnWidths) {
        int[] scaled = new int[columnWidths.length];
        for (int i = 0; i < columnWidths.length; i++) {
            scaled[i] = scaleWidth(columnWidths[i]);
        }
        return scaled;
    }

    private static int scaleWidth(int width) {
        return Math.max(1, (int) Math.round(width * REQUEST_TABLE_WIDTH_SCALE));
    }

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthDxa));
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize, boolean bold,
                                         ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);
        run.setBold(bold);
    }

    private static String normalizeSpace(String value) {
        if (value == null) {
            return "";
        }
        return value.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
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
}
