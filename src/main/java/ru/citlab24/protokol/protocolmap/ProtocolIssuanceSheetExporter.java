package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Locale;

final class ProtocolIssuanceSheetExporter {
    private static final int MAP_PROTOCOL_NUMBER_ROW_INDEX = 21;
    private static final int MAP_CUSTOMER_ROW_INDEX = 5;
    private static final int MAP_APPLICATION_ROW_INDEX = 22;
    private static final String ISSUANCE_SHEET_NAME = "лист выдачи протоколов.docx";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;

    private ProtocolIssuanceSheetExporter() {
    }

    static void generate(File sourceFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = new File(mapFile.getParentFile(), ISSUANCE_SHEET_NAME);
        String protocolNumber = resolveProtocolNumberFromMap(mapFile);
        String protocolDate = resolveProtocolDateFromSource(sourceFile);
        String customerName = resolveCustomerNameFromMap(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Лист выдачи протоколов");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(FONT_SIZE);

            XWPFTable table = document.createTable(2, 5);
            setTableCellText(table.getRow(0).getCell(0), "№ п/п");
            setTableCellText(table.getRow(0).getCell(1), "Номер протокола");
            setTableCellText(table.getRow(0).getCell(2), "Дата протокола");
            setTableCellText(table.getRow(0).getCell(3),
                    "Наименование заказчика (полное или сокращенное для юридического лица, " +
                            "ФИО (допустимо указание только фамилии) для физического лица)");
            setTableCellText(table.getRow(0).getCell(4), "Номер заявки");

            setTableCellText(table.getRow(1).getCell(0), "1");
            setTableCellText(table.getRow(1).getCell(1), protocolNumber);
            setTableCellText(table.getRow(1).getCell(2), protocolDate);
            setTableCellText(table.getRow(1).getCell(3), customerName);
            setTableCellText(table.getRow(1).getCell(4), applicationNumber);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (IOException ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    private static void setTableCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
    }

    private static String resolveProtocolNumberFromMap(File mapFile) {
        String line = readMapRowText(mapFile, MAP_PROTOCOL_NUMBER_ROW_INDEX);
        return extractValueAfterPrefix(line, "1. Номер протокола");
    }

    private static String resolveCustomerNameFromMap(File mapFile) {
        String line = readMapRowText(mapFile, MAP_CUSTOMER_ROW_INDEX);
        String value = extractValueAfterPrefix(line, "1. Заказчик");
        if (value.isBlank()) {
            value = extractValueAfterPrefix(line, "Заказчик");
        }
        int commaIndex = value.indexOf(',');
        if (commaIndex >= 0) {
            value = value.substring(0, commaIndex).trim();
        }
        return value;
    }

    private static String resolveApplicationNumberFromMap(File mapFile) {
        String line = readMapRowText(mapFile, MAP_APPLICATION_ROW_INDEX);
        String lowerLine = line.toLowerCase(Locale.ROOT);
        int applicationIndex = lowerLine.indexOf("заявка");
        if (applicationIndex >= 0) {
            String value = line.substring(applicationIndex + "заявка".length()).trim();
            return trimLeadingPunctuation(value);
        }
        int commaIndex = line.indexOf(',');
        if (commaIndex >= 0 && commaIndex + 1 < line.length()) {
            return line.substring(commaIndex + 1).trim();
        }
        return line.trim();
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

    private static String resolveProtocolDateFromSource(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        String name = sourceFile.getName().toLowerCase(Locale.ROOT);
        if (!name.endsWith(".docx")) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             XWPFDocument document = new XWPFDocument(in)) {
            java.util.List<String> lines = extractDocLines(document);
            java.util.List<String> meaningfulLines = new ArrayList<>();
            for (String line : lines) {
                if (line != null && !line.isBlank()) {
                    meaningfulLines.add(line.trim());
                }
            }
            if (meaningfulLines.size() >= 7) {
                return meaningfulLines.get(6);
            }
        } catch (IOException ignored) {
            return "";
        }
        return "";
    }

    private static java.util.List<String> extractDocLines(XWPFDocument document) {
        java.util.List<String> lines = new ArrayList<>();
        for (org.apache.poi.xwpf.usermodel.IBodyElement element : document.getBodyElements()) {
            if (element instanceof org.apache.poi.xwpf.usermodel.XWPFParagraph paragraph) {
                addDocLine(lines, paragraph.getText());
            } else if (element instanceof org.apache.poi.xwpf.usermodel.XWPFTable table) {
                for (org.apache.poi.xwpf.usermodel.XWPFTableRow row : table.getRows()) {
                    StringBuilder builder = new StringBuilder();
                    for (XWPFTableCell cell : row.getTableCells()) {
                        String cellText = cell.getText();
                        if (cellText != null && !cellText.isBlank()) {
                            if (builder.length() > 0) {
                                builder.append(' ');
                            }
                            builder.append(cellText.trim());
                        }
                    }
                    if (builder.length() > 0) {
                        addDocLine(lines, builder.toString());
                    }
                }
            }
        }
        return lines;
    }

    private static void addDocLine(java.util.List<String> lines, String text) {
        if (text == null) {
            return;
        }
        String[] split = text.split("\\R");
        for (String line : split) {
            lines.add(line);
        }
    }

    private static String readMapRowText(File mapFile, int rowIndex) {
        if (mapFile == null || !mapFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            String firstValue = "";
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.isEmpty()) {
                    continue;
                }
                if (firstValue.isEmpty()) {
                    firstValue = text;
                }
                if (text.contains("Номер протокола")
                        || text.contains("Заказчик")
                        || text.toLowerCase(Locale.ROOT).contains("договор")
                        || text.toLowerCase(Locale.ROOT).contains("заявка")) {
                    return text;
                }
            }
            return firstValue;
        } catch (Exception ex) {
            return "";
        }
    }

    private static String extractValueAfterPrefix(String line, String prefix) {
        if (line == null) {
            return "";
        }
        String normalized = line.trim();
        int index = normalized.indexOf(prefix);
        if (index >= 0) {
            return normalized.substring(index + prefix.length()).replace(":", "").trim();
        }
        return normalized.replace(":", "").trim();
    }
}
