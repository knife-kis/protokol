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
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

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
    private static final String NORMATIVE_REQUIREMENTS_WITH_BRACKET_MARKER = "(нормативные требования";
    private static final String LNW_SENTENCE_MARKER =
            "Индекс приведенного уровня ударного шума (Lnw) по результатам измерений для перекрытия между помещениями";
    private static final String RW_FLOOR_SENTENCE_MARKER =
            "Индекс изоляции воздушного шума (Rw) по результатам измерений для перекрытия между помещениями";
    private static final String RW_PARTITION_SENTENCE_MARKER =
            "Индекс изоляции воздушного шума (Rw) по результатам измерений для перегородки между помещениями";
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
        if (!protocolFile.getName().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            return result;
        }

        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {

            List<String> lnRequirements = extractLnRequirements(document);

            List<String> rwFloorRequirements = extractRRequirements(
                    document,
                    RW_FLOOR_SENTENCE_MARKER,
                    "Для перекрытия между помещениями "
            );

            List<String> rwPartitionRequirements = extractRRequirements(
                    document,
                    RW_PARTITION_SENTENCE_MARKER,
                    "Для перегородки между помещениями "
            );

            appendGroup(result, lnRequirements);
            appendGroup(result, rwFloorRequirements);
            appendGroup(result, rwPartitionRequirements);

        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }

        return result;
    }

    private static void appendGroup(List<String> target, List<String> group) {
        if (group.isEmpty()) {
            return;
        }
        if (!target.isEmpty()) {
            target.add("");
        }
        target.addAll(group);
    }

    private static List<String> extractLnRequirements(XWPFDocument document) {
        List<String> result = new ArrayList<>();
        if (document == null) {
            return result;
        }

        List<IBodyElement> elements = document.getBodyElements();

        int start = findSectionStartIndex(elements, "17.1", 0);
        if (start < 0) {
            return result;
        }

        int end = findSectionStartIndex(elements, "17.2", start + 1);
        if (end < 0) {
            end = elements.size();
        }

        // Логика как ты описал: в 17.1 ищем таблицу где есть "Ln, дБ",
        // затем берём абзац сразу после этой таблицы и парсим в нём LNW_SENTENCE_MARKER + (нормативные требования ...)
        for (int i = start; i < end; i++) {
            IBodyElement element = elements.get(i);
            if (!(element instanceof XWPFTable table)) {
                continue;
            }
            if (!tableContainsLnDbMarker(table)) {
                continue;
            }

            String nextParagraphText = findNextNonEmptyParagraphText(elements, i + 1, end);
            if (nextParagraphText.isBlank()) {
                continue;
            }

            result.addAll(extractRequirementLinesFromText(
                    nextParagraphText,
                    LNW_SENTENCE_MARKER,
                    "Для перекрытия между помещениями "
            ));
        }

        return distinctPreserveOrder(result);
    }


    private static List<String> extractRRequirements(XWPFDocument document, String sentenceMarker, String prefix) {
        List<String> result = new ArrayList<>();
        if (document == null || sentenceMarker == null || sentenceMarker.isBlank()) {
            return result;
        }

        List<IBodyElement> elements = document.getBodyElements();

        int start = findSectionStartIndex(elements, "17.2", 0);
        if (start < 0) {
            return result;
        }

        // Важно: конец секции ищем как ЗАГОЛОВОК "18." (в начале абзаца), а не просто contains("18.")
        int end = findSectionStartIndex(elements, "18.", start + 1);
        if (end < 0) {
            end = elements.size();
        }

        for (int i = start; i < end; i++) {
            IBodyElement element = elements.get(i);

            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (text.isBlank()) {
                    continue;
                }
                result.addAll(extractRequirementLinesFromText(text, sentenceMarker, prefix));
                continue;
            }

            // На всякий случай: если фраза почему-то попала в таблицу
            if (element instanceof XWPFTable table) {
                for (XWPFTableRow row : table.getRows()) {
                    String rowText = tableRowToText(row);
                    if (rowText.isBlank()) {
                        continue;
                    }
                    result.addAll(extractRequirementLinesFromText(rowText, sentenceMarker, prefix));
                }
            }
        }

        return distinctPreserveOrder(result);
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

    private static String extractNormativeRequirements(String source) {
        if (source == null || source.isBlank()) {
            return "";
        }
        Pattern pattern = Pattern.compile("\\(\\s*нормативные требования\\s*([^)]*)\\)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
        Matcher matcher = pattern.matcher(source);
        if (!matcher.find()) {
            return "";
        }
        return normalizeSpace(matcher.group(1));
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
    private static int findSectionStartIndex(List<IBodyElement> elements, String sectionNumber, int fromIndex) {
        if (elements == null || elements.isEmpty() || sectionNumber == null || sectionNumber.isBlank()) {
            return -1;
        }

        // Ищем заголовок вида "17.1." / "17.1 " / "18." и т.п.
        // Важно: НЕ должны матчиться подзаголовки типа "17.1.1"
        String base = normalizeSpace(sectionNumber);
        if (base.endsWith(".")) {
            base = base.substring(0, base.length() - 1).trim();
        }
        if (base.isBlank()) {
            return -1;
        }

        Pattern headingPattern = Pattern.compile(
                "^\\s*" + Pattern.quote(base) + "\\s*\\.?\\s*(?!\\d)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
        );

        int start = Math.max(0, fromIndex);
        for (int i = start; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (headingPattern.matcher(text).find()) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static String findNextNonEmptyParagraphText(List<IBodyElement> elements, int fromIndex, int endExclusive) {
        if (elements == null || elements.isEmpty()) {
            return "";
        }
        int to = Math.min(endExclusive, elements.size());
        for (int i = Math.max(0, fromIndex); i < to; i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (!text.isBlank()) {
                    return text;
                }
            }
        }
        return "";
    }

    private static boolean tableContainsLnDbMarker(XWPFTable table) {
        if (table == null) {
            return false;
        }
        Pattern lnPattern = Pattern.compile("Ln\\s*,\\s*дБ", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                String text = normalizeSpace(cell.getText());
                if (lnPattern.matcher(text).find()) {
                    return true;
                }
            }
        }
        return false;
    }

    private static String tableRowToText(XWPFTableRow row) {
        if (row == null) {
            return "";
        }
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
        return builder.toString().trim();
    }

    private static List<String> extractRequirementLinesFromText(String text, String sentenceMarker, String prefix) {
        List<String> result = new ArrayList<>();
        if (text == null || text.isBlank() || sentenceMarker == null || sentenceMarker.isBlank()) {
            return result;
        }

        String normalized = normalizeSpace(text);

        // Вытаскиваем строго "A и B" после маркера и норму из "(нормативные требования ...)"
        Pattern pattern = Pattern.compile(
                Pattern.quote(sentenceMarker)
                        + "\\s*([A-Za-zА-Яа-я0-9]+)\\s+и\\s+([A-Za-zА-Яа-я0-9]+).*?"
                        + "\\(\\s*нормативные требования\\s*([^)]*)\\)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
        );

        Matcher matcher = pattern.matcher(normalized);
        while (matcher.find()) {
            String first = normalizeSpace(matcher.group(1));
            String second = normalizeSpace(matcher.group(2));
            String norm = normalizeSpace(matcher.group(3));

            if (first.isBlank() || second.isBlank() || norm.isBlank()) {
                continue;
            }

            String rooms = first + " и " + second;
            String line = (prefix == null ? "" : prefix) + rooms + " (нормативные требования " + norm + ")";
            result.add(normalizeSpace(line));
        }

        return result;
    }

    private static List<String> distinctPreserveOrder(List<String> values) {
        java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<>();
        if (values != null) {
            for (String value : values) {
                String v = normalizeSpace(value);
                if (!v.isBlank()) {
                    set.add(v);
                }
            }
        }
        return new ArrayList<>(set);
    }

}
