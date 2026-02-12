package ru.citlab24.protokol.tabs.qms;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableRowAlign;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

final class VlkEquipmentIssuanceWordExporter {

    private static final int FONT_SIZE = 9;

    private VlkEquipmentIssuanceWordExporter() {
    }

    static void export(File target, VlkWordExporter.PlanRow row) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            configurePortraitPage(document);
            createHeader(document);
            addCenterBlock(document, row);
            addMainTable(document);

            try (FileOutputStream out = new FileOutputStream(target)) {
                document.write(out);
            }
        }
    }

    private static void configurePortraitPage(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                ? document.getDocument().getBody().getSectPr()
                : document.getDocument().getBody().addNewSectPr();

        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(11906));
        pageSize.setH(BigInteger.valueOf(16838));

        CTPageMar margins = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();
        margins.setTop(BigInteger.valueOf(720));
        margins.setRight(BigInteger.valueOf(720));
        margins.setBottom(BigInteger.valueOf(720));
        margins.setLeft(BigInteger.valueOf(720));
        margins.setHeader(BigInteger.valueOf(540));
    }

    private static void createHeader(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                ? document.getDocument().getBody().getSectPr()
                : document.getDocument().getBody().addNewSectPr();

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        XWPFTable headerTable = header.createTable(3, 3);
        headerTable.setTableAlignment(TableRowAlign.CENTER);
        headerTable.setWidth("100%");
        setFixedLayout(headerTable);
        setGrid(headerTable, 3200, 3200, 3200);

        setCellText(headerTable.getRow(0).getCell(0),
                "Испытательная лаборатория ООО «ЦИТ»", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(headerTable.getRow(0).getCell(1),
                "Лист выдачи приборов Ф39 ДП ИЛ 2-2023", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(headerTable.getRow(0).getCell(2),
                "Дата утверждения бланка формуляра: 01.01.2023г.", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(headerTable.getRow(1).getCell(2), "Редакция № 1", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setPageCounterCellText(headerTable.getRow(2).getCell(2));

        mergeCellsVertically(headerTable, 0, 0, 2);
        mergeCellsVertically(headerTable, 1, 0, 2);
        styleAllCellParagraphs(headerTable, FONT_SIZE);
    }

    private static void addCenterBlock(XWPFDocument document, VlkWordExporter.PlanRow row) {
        addCenteredParagraph(document,
                "Место использования прибора:\nКонтроль точности результатов измерений " + methodCipher(row.event()),
                true);
        addCenteredParagraph(document,
                "Фамилия, инициалы лица, получившего приборы:\n" + (row.responsible() == null ? "" : row.responsible()),
                true);
        addCenteredParagraph(document, "Лист выдачи приборов", true);
    }

    private static void addMainTable(XWPFDocument document) {
        XWPFTable table = document.createTable(2, 5);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 1700, 1200, 2600, 2600, 1500);

        setCellText(table.getRow(0).getCell(0), "Наименование\nоборудования", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(table.getRow(0).getCell(1), "Зав. №", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(table.getRow(0).getCell(2),
                "Получение прибора, отметка о проведении технического обслуживания и (или) проверки перед выдачей",
                ParagraphAlignment.CENTER,
                true,
                FONT_SIZE);
        setCellText(table.getRow(0).getCell(3),
                "Возврат оборудования, Отметка о проведении технического обслуживания и (или) проверки перед возвратом оборудования и после его транспортировки",
                ParagraphAlignment.CENTER,
                true,
                FONT_SIZE);
        setCellText(table.getRow(0).getCell(4), "Примечание", ParagraphAlignment.CENTER, true, FONT_SIZE);

        styleAllCellParagraphs(table, FONT_SIZE);
    }

    private static String methodCipher(String event) {
        String prefix = "Контроль точности результатов измерений по";
        if (event == null) {
            return "";
        }
        if (event.startsWith(prefix)) {
            return event.substring(prefix.length()).trim();
        }
        return event;
    }

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(rowIndex == fromRow ? STMerge.RESTART : STMerge.CONTINUE);
            if (rowIndex > fromRow) {
                setCellText(cell, "", ParagraphAlignment.LEFT, false, FONT_SIZE);
            }
        }
    }

    private static void setPageCounterCellText(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun textRun = paragraph.createRun();
        textRun.setFontFamily("Arial");
        textRun.setFontSize(FONT_SIZE);
        textRun.setText("Количество страниц: ");

        addField(paragraph, "PAGE");

        XWPFRun separatorRun = paragraph.createRun();
        separatorRun.setFontFamily("Arial");
        separatorRun.setFontSize(FONT_SIZE);
        separatorRun.setText(" / ");

        addField(paragraph, "NUMPAGES");
    }

    private static void addField(XWPFParagraph paragraph, String fieldName) {
        XWPFRun beginRun = paragraph.createRun();
        beginRun.setFontFamily("Arial");
        beginRun.setFontSize(FONT_SIZE);
        beginRun.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun instrRun = paragraph.createRun();
        instrRun.setFontFamily("Arial");
        instrRun.setFontSize(FONT_SIZE);
        instrRun.getCTR().addNewInstrText().setStringValue(fieldName);

        XWPFRun separateRun = paragraph.createRun();
        separateRun.setFontFamily("Arial");
        separateRun.setFontSize(FONT_SIZE);
        separateRun.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun valueRun = paragraph.createRun();
        valueRun.setFontFamily("Arial");
        valueRun.setFontSize(FONT_SIZE);
        valueRun.setText("1");

        XWPFRun endRun = paragraph.createRun();
        endRun.setFontFamily("Arial");
        endRun.setFontSize(FONT_SIZE);
        endRun.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void setFixedLayout(XWPFTable table) {
        CTTbl cttbl = table.getCTTbl();
        CTTblPr tblPr = cttbl.getTblPr();
        if (tblPr == null) {
            tblPr = cttbl.addNewTblPr();
        }
        CTTblLayoutType layout = tblPr.isSetTblLayout() ? tblPr.getTblLayout() : tblPr.addNewTblLayout();
        layout.setType(org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType.FIXED);
    }

    private static void setGrid(XWPFTable table, int... widths) {
        CTTbl cttbl = table.getCTTbl();
        CTTblGrid tblGrid = cttbl.getTblGrid();
        if (tblGrid == null) {
            tblGrid = cttbl.addNewTblGrid();
        } else {
            while (tblGrid.sizeOfGridColArray() > 0) {
                tblGrid.removeGridCol(0);
            }
        }

        for (int width : widths) {
            CTTblGridCol col = tblGrid.addNewGridCol();
            col.setW(BigInteger.valueOf(width));
        }
    }

    private static void setCellText(XWPFTableCell cell,
                                    String text,
                                    ParagraphAlignment alignment,
                                    boolean bold,
                                    int size) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(size);
        run.setBold(bold);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    private static void addCenteredParagraph(XWPFDocument document, String text, boolean bold) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(FONT_SIZE);
        run.setBold(bold);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    private static void styleAllCellParagraphs(XWPFTable table, int size) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    if (paragraph.getRuns() == null) {
                        continue;
                    }
                    for (XWPFRun run : paragraph.getRuns()) {
                        run.setFontFamily("Arial");
                        run.setFontSize(size);
                    }
                }
            }
        }
    }
}
