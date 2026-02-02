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

    private SoundInsulationRequestFormExporter() {
    }

    static void generate(File protocolFile, File mapFile, String workDeadline, String customerInn) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveRequestFormFile(mapFile);
        String applicationNumber = RequestFormExporter.resolveApplicationNumberFromMap(mapFile);
        File planFile = SoundInsulationMeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
        List<List<String>> planRows = extractPlanRows(planFile);
        List<String> customerLines = extractCustomerDataLines(protocolFile);

        if (planRows.isEmpty()) {
            List<String> empty = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                empty.add("");
            }
            planRows.add(empty);
        }

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

            XWPFParagraph pageBreak = document.createParagraph();
            setParagraphSpacing(pageBreak);
            pageBreak.createRun().addBreak(org.apache.poi.xwpf.usermodel.BreakType.PAGE);

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

            XWPFParagraph highlightParagraph = document.createParagraph();
            setParagraphSpacing(highlightParagraph);
            XWPFRun highlightRun = highlightParagraph.createRun();
            highlightRun.setFontFamily(FONT_NAME);
            highlightRun.setFontSize(FONT_SIZE);
            highlightRun.setText("Для перекрытия между помещениями");
            highlightRun.setTextHighlightColor("yellow");

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
            lines.addAll(extractSection(allLines, ROOM_PARAMS_START, ROOM_PARAMS_END));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return lines;
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
}
