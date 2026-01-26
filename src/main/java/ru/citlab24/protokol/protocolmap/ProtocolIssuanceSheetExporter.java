package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;

import java.math.BigInteger;


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
        File targetFile = new File(mapFile.getParentFile(), ISSUANCE_SHEET_NAME);
        String protocolNumber = resolveProtocolNumberFromMap(mapFile);
        String protocolDate = resolveProtocolDateFromSource(sourceFile);
        String customerName = resolveCustomerNameFromMap(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            setLandscapeOrientation(document);

            // НОВОЕ: колонтитул как в шаблоне
            applyStandardHeader(document);

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
            java.math.BigInteger width = (java.math.BigInteger) pageSize.getW();
            pageSize.setW(pageSize.getH());
            pageSize.setH(width);
        } else {
            pageSize.setW(java.math.BigInteger.valueOf(16840));
            pageSize.setH(java.math.BigInteger.valueOf(11900));
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
    private static void applyStandardHeader(XWPFDocument document) {
        // Гарантируем секцию
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        // В Word у header обычно есть 1 пустой параграф — оставляем его, как в шаблоне.
        XWPFTable tbl = header.createTable(3, 3);
        configureHeaderTableLikeTemplate(tbl);

        // Тексты как в файле-шаблоне :contentReference[oaicite:1]{index=1}
        String left = "Испытательная лаборатория ООО «ЦИТ»";
        String center = "Лист выдачи протоколов Ф17 ДП ИЛ 2-2023";

        setHeaderCellText(tbl.getRow(0).getCell(0), left);
        setHeaderCellText(tbl.getRow(1).getCell(0), left);
        setHeaderCellText(tbl.getRow(2).getCell(0), left);

        setHeaderCellText(tbl.getRow(0).getCell(1), center);
        setHeaderCellText(tbl.getRow(1).getCell(1), center);
        setHeaderCellText(tbl.getRow(2).getCell(1), center);

        setHeaderCellText(tbl.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(tbl.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellPageCount(tbl.getRow(2).getCell(2));
    }

    private static void configureHeaderTableLikeTemplate(XWPFTable table) {
        // Таблица 3×3, fixed layout, centered, borders single sz=4,
        // ширина 9639 и gridCol: 2951 / 3140 / 3548.
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(9639));

        // ВАЖНО: для таблицы это CTJcTable, НЕ CTJc
        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(STJcTable.CENTER);

        CTTblLayoutType layout = pr.isSetTblLayout() ? pr.getTblLayout() : pr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);

        CTTblBorders borders = pr.isSetTblBorders() ? pr.getTblBorders() : pr.addNewTblBorders();
        setBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        setBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        setBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        setBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        setBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        setBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());

        // grid (без unsetTblGrid — его может не быть)
        CTTblGrid grid = ct.getTblGrid();
        if (grid == null) {
            grid = ct.addNewTblGrid();
        } else {
            while (grid.sizeOfGridColArray() > 0) {
                grid.removeGridCol(0);
            }
        }
        grid.addNewGridCol().setW(BigInteger.valueOf(2951));
        grid.addNewGridCol().setW(BigInteger.valueOf(3140));
        grid.addNewGridCol().setW(BigInteger.valueOf(3548));

        // ширины ячеек (для fixed layout)
        setCellWidth(table, 0, 0, 2951);
        setCellWidth(table, 1, 0, 2951);
        setCellWidth(table, 2, 0, 2951);

        setCellWidth(table, 0, 1, 3140);
        setCellWidth(table, 1, 1, 3140);
        setCellWidth(table, 2, 1, 3140);

        setCellWidth(table, 0, 2, 3548);
        setCellWidth(table, 1, 2, 3548);
        setCellWidth(table, 2, 2, 3548);
    }

    private static void setBorder(CTBorder b) {
        b.setVal(STBorder.SINGLE);
        b.setSz(BigInteger.valueOf(4));
        b.setSpace(BigInteger.ZERO);
        b.setColor("auto");
    }

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth w = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        w.setType(STTblWidth.DXA);
        w.setW(BigInteger.valueOf(widthDxa));
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun r = p.createRun();
        r.setText(text != null ? text : "");
        r.setFontFamily(FONT_NAME);
        r.setFontSize(FONT_SIZE);
    }

    private static void setHeaderCellPageCount(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);

        XWPFRun r0 = p.createRun();
        r0.setText("Количество страниц: ");
        r0.setFontFamily(FONT_NAME);
        r0.setFontSize(FONT_SIZE);

        appendField(p, "PAGE");
        XWPFRun rSep = p.createRun();
        rSep.setText(" / ");
        rSep.setFontFamily(FONT_NAME);
        rSep.setFontSize(FONT_SIZE);

        appendField(p, "NUMPAGES");
    }

    private static void appendField(XWPFParagraph p, String instr) {
        XWPFRun rBegin = p.createRun();
        rBegin.setFontFamily(FONT_NAME);
        rBegin.setFontSize(FONT_SIZE);
        rBegin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun rInstr = p.createRun();
        rInstr.setFontFamily(FONT_NAME);
        rInstr.setFontSize(FONT_SIZE);
        var ctText = rInstr.getCTR().addNewInstrText();
        ctText.setStringValue(instr);
        ctText.setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFRun rSep = p.createRun();
        rSep.setFontFamily(FONT_NAME);
        rSep.setFontSize(FONT_SIZE);
        rSep.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        // плейсхолдер — Word при открытии обновит поле
        XWPFRun rText = p.createRun();
        rText.setFontFamily(FONT_NAME);
        rText.setFontSize(FONT_SIZE);
        rText.setText("1");

        XWPFRun rEnd = p.createRun();
        rEnd.setFontFamily(FONT_NAME);
        rEnd.setFontSize(FONT_SIZE);
        rEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

}
