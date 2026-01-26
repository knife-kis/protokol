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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
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
        File targetFile = resolveIssuanceSheetFile(mapFile);
        String protocolNumber = resolveProtocolNumberFromMap(mapFile);
        String protocolDate = resolveProtocolDateFromSource(sourceFile);
        String customerName = resolveCustomerNameFromMap(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            setLandscapeOrientation(document);
            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Лист выдачи протоколов");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(FONT_SIZE);

            XWPFTable table = document.createTable(4, 5);
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
            setTableCellText(table.getRow(2).getCell(2), "");
            setTableCellText(table.getRow(2).getCell(3), "");
            setTableCellText(table.getRow(2).getCell(4), "");
            setTableCellText(table.getRow(3).getCell(2), "");
            setTableCellText(table.getRow(3).getCell(3), "");
            setTableCellText(table.getRow(3).getCell(4), "");

            mergeCellsVertically(table, 0, 1, 3);
            mergeCellsVertically(table, 1, 1, 3);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (IOException ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    private static void setLandscapeOrientation(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }
        CTPageSz pageSize = sectPr.getPgSz();
        if (pageSize == null) {
            pageSize = sectPr.addNewPgSz();
        }
        pageSize.setOrient(STPageOrientation.LANDSCAPE);
        if (pageSize.isSetW() && pageSize.isSetH()) {
            java.math.BigInteger width = pageSize.getW();
            pageSize.setW(pageSize.getH());
            pageSize.setH(width);
        } else {
            pageSize.setW(java.math.BigInteger.valueOf(16840));
            pageSize.setH(java.math.BigInteger.valueOf(11900));
        }
    }

    static File resolveIssuanceSheetFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), ISSUANCE_SHEET_NAME);
    }

    private static void setTableCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
    }

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (cell == null) {
                continue;
            }
            if (cell.getCTTc().getTcPr() == null) {
                cell.getCTTc().addNewTcPr();
            }
            if (rowIndex == fromRow) {
                cell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
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
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(6);
            if (row == null) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (!text.isEmpty()) {
                    return text;
                }
            }
        } catch (Exception ignored) {
            return "";
        }
        return "";
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
