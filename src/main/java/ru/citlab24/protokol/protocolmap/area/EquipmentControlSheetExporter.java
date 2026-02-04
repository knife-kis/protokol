package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblCellMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

final class EquipmentControlSheetExporter {
    private static final String CONTROL_SHEET_COAL_NAME = "Лист контроля оборудования Уголь.docx";
    private static final String CONTROL_SHEET_SOURCE_NAME = "Лист контроля оборудования Источник.docx";
    private static final String CONTROL_SHEET_TITLE = "Лист контроля оборудования";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int TABLE_WIDTH = 12560;

    private EquipmentControlSheetExporter() {
    }

    static void generate(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        generateControlSheet(
                mapFile,
                CONTROL_SHEET_COAL_NAME,
                CONTROL_SHEET_TITLE,
                "Камера-01",
                "631",
                "Измерения фона активированного угля",
                new String[][]{
                        {"", "БДБ-2009", "", "Выбираем среднее значение", ""},
                        {"", "БДБ-2010", "", "Выбираем среднее значение", ""}
                }
        );
        generateControlSheet(
                mapFile,
                CONTROL_SHEET_SOURCE_NAME,
                CONTROL_SHEET_TITLE,
                "Камера-01",
                "631",
                "измерения по контрольному источнику №647",
                new String[][]{
                        {"", "БДБ-2009", "от 93 имп/с до 123 имп/с", "", ""},
                        {"", "БДБ-2010", "от 93 имп/с до 123 имп/с", "", ""}
                }
        );
    }

    static File resolveControlSheetFile(File mapFile, String fileName) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), fileName);
    }

    static List<File> resolveControlSheetFiles(File mapFile) {
        if (mapFile == null) {
            return List.of();
        }
        return List.of(
                resolveControlSheetFile(mapFile, CONTROL_SHEET_COAL_NAME),
                resolveControlSheetFile(mapFile, CONTROL_SHEET_SOURCE_NAME)
        );
    }

    private static void generateControlSheet(
            File mapFile,
            String fileName,
            String controlSheetTitle,
            String equipmentName,
            String inventoryNumber,
            String controlNotes,
            String[][] controlRows
    ) {
        File targetFile = resolveControlSheetFile(mapFile, fileName);

        try (XWPFDocument document = new XWPFDocument()) {
            applyControlSheetHeader(document, controlSheetTitle);
            addControlSheetSection(
                    document,
                    controlSheetTitle,
                    equipmentName,
                    inventoryNumber,
                    controlNotes,
                    controlRows,
                    4
            );

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    private static void applyControlSheetHeader(XWPFDocument document, String controlSheetTitle) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }
        applyLandscapePageSettings(sectPr);
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        XWPFTable table = header.createTable(3, 3);
        configureHeaderTableLikeTemplate(table);

        setHeaderCellText(table.getRow(0).getCell(0), "Испытательная лаборатория ООО «ЦИТ»");
        setHeaderCellText(table.getRow(0).getCell(1), CONTROL_SHEET_TITLE + "\nФ1 РИ ИЛ 2-2023");
        setHeaderCellText(table.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(table.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellPageCount(table.getRow(2).getCell(2));

        mergeCellsVertically(table, 0, 0, 2);
        mergeCellsVertically(table, 1, 0, 2);
    }

    private static void addControlSheetSection(
            XWPFDocument document,
            String controlSheetTitle,
            String equipmentName,
            String inventoryNumber,
            String controlNotes,
            String[][] controlRows,
            int rowHeightBreaks
    ) {
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(title);
        XWPFRun titleRun = title.createRun();
        titleRun.setText(controlSheetTitle);
        titleRun.setFontFamily(FONT_NAME);
        titleRun.setFontSize(FONT_SIZE);
        titleRun.setBold(true);

        XWPFTable equipmentTable = document.createTable(2, 3);
        configureTableLayout(equipmentTable, new int[]{4187, 4187, 4186});
        setTableCellText(equipmentTable.getRow(0).getCell(0), "Наименование оборудования",
                FONT_SIZE, true, ParagraphAlignment.CENTER, 0);
        setTableCellText(equipmentTable.getRow(0).getCell(1), "Инв. номер оборудования",
                FONT_SIZE, true, ParagraphAlignment.CENTER, 0);
        setTableCellText(equipmentTable.getRow(0).getCell(2), "Указания к контролю",
                FONT_SIZE, true, ParagraphAlignment.CENTER, 0);
        setTableCellText(equipmentTable.getRow(1).getCell(0), equipmentName,
                FONT_SIZE, false, ParagraphAlignment.CENTER, 0);
        setTableCellText(equipmentTable.getRow(1).getCell(1), inventoryNumber,
                FONT_SIZE, false, ParagraphAlignment.CENTER, 0);
        setTableCellText(equipmentTable.getRow(1).getCell(2), controlNotes,
                FONT_SIZE, false, ParagraphAlignment.CENTER, 0);

        addSpacer(document);

        XWPFTable controlTable = document.createTable(3, 5);
        configureTableLayout(controlTable, new int[]{2323, 2323, 2323, 2323, 3268});
        String[] headerCells = new String[]{
                "Дата проведения контроля",
                "Результат контроля, ед.из.",
                "Допустимое отклонение, ед.из.",
                "Примечание",
                "Не выходит за установленные пределы (+)/выходит за установленные пределы " +
                        "(-), подпись ответственного"
        };
        for (int colIndex = 0; colIndex < headerCells.length; colIndex++) {
            setTableCellText(controlTable.getRow(0).getCell(colIndex), headerCells[colIndex],
                    FONT_SIZE, true, ParagraphAlignment.CENTER, 0);
        }

        for (int rowIndex = 0; rowIndex < controlRows.length; rowIndex++) {
            XWPFTableRow row = controlTable.getRow(rowIndex + 1);
            for (int colIndex = 0; colIndex < controlRows[rowIndex].length; colIndex++) {
                setTableCellText(row.getCell(colIndex), controlRows[rowIndex][colIndex],
                        FONT_SIZE, false, ParagraphAlignment.CENTER, rowHeightBreaks);
            }
        }

    }

    private static void addSpacer(XWPFDocument document) {
        XWPFParagraph spacer = document.createParagraph();
        setParagraphSpacing(spacer);
        spacer.createRun().addBreak();
    }

    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph pageBreak = document.createParagraph();
        setParagraphSpacing(pageBreak);
        XWPFRun run = pageBreak.createRun();
        run.addBreak(BreakType.PAGE);
    }

    private static void applyLandscapePageSettings(CTSectPr sectPr) {
        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setOrient(STPageOrientation.LANDSCAPE);
        pageSize.setW(BigInteger.valueOf(16838));
        pageSize.setH(BigInteger.valueOf(11906));
    }

    private static void configureHeaderTableLikeTemplate(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(9639));

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
        applyMinimalHeaderCellMargins(pr);

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

    private static void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
        border.setColor("auto");
    }

    private static void applyMinimalHeaderCellMargins(CTTblPr pr) {
        CTTblCellMar cellMar = pr.isSetTblCellMar() ? pr.getTblCellMar() : pr.addNewTblCellMar();
        setCellMargin(cellMar.isSetTop() ? cellMar.getTop() : cellMar.addNewTop());
        setCellMargin(cellMar.isSetLeft() ? cellMar.getLeft() : cellMar.addNewLeft());
        setCellMargin(cellMar.isSetBottom() ? cellMar.getBottom() : cellMar.addNewBottom());
        setCellMargin(cellMar.isSetRight() ? cellMar.getRight() : cellMar.addNewRight());
    }

    private static void setCellMargin(CTTblWidth margin) {
        margin.setType(STTblWidth.DXA);
        margin.setW(BigInteger.ZERO);
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(paragraph);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            XWPFRun run = paragraph.createRun();
            run.setText(lines[i]);
            run.setFontFamily(FONT_NAME);
            run.setFontSize(FONT_SIZE);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setHeaderCellPageCount(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(paragraph);

        XWPFRun run = paragraph.createRun();
        run.setText("Количество страниц: ");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);

        XWPFRun fieldBegin = paragraph.createRun();
        fieldBegin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun fieldCode = paragraph.createRun();
        fieldCode.getCTR().addNewInstrText().setStringValue("PAGE ");
        fieldCode.getCTR().getInstrTextArray(0).setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFRun fieldSeparator = paragraph.createRun();
        fieldSeparator.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun fieldValue = paragraph.createRun();
        fieldValue.setText("1");
        fieldValue.setFontFamily(FONT_NAME);
        fieldValue.setFontSize(FONT_SIZE);

        XWPFRun fieldEnd = paragraph.createRun();
        fieldEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

        XWPFRun runTotal = paragraph.createRun();
        runTotal.setText(" / ");
        runTotal.setFontFamily(FONT_NAME);
        runTotal.setFontSize(FONT_SIZE);

        XWPFRun fieldBeginTotal = paragraph.createRun();
        fieldBeginTotal.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun fieldCodeTotal = paragraph.createRun();
        fieldCodeTotal.getCTR().addNewInstrText().setStringValue("NUMPAGES ");
        fieldCodeTotal.getCTR().getInstrTextArray(0).setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFRun fieldSeparatorTotal = paragraph.createRun();
        fieldSeparatorTotal.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun fieldValueTotal = paragraph.createRun();
        fieldValueTotal.setText("1");
        fieldValueTotal.setFontFamily(FONT_NAME);
        fieldValueTotal.setFontSize(FONT_SIZE);

        XWPFRun fieldEndTotal = paragraph.createRun();
        fieldEndTotal.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void mergeCellsVertically(XWPFTable table, int column, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(column);
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

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(TABLE_WIDTH));

        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(STJcTable.CENTER);

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
        for (int width : columnWidths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }

        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < columnWidths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, columnWidths[colIndex]);
            }
        }
    }

    private static void setCellWidth(XWPFTable table, int rowIndex, int colIndex, int width) {
        XWPFTableCell cell = table.getRow(rowIndex).getCell(colIndex);
        if (cell == null) {
            return;
        }
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth tcW = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        tcW.setType(STTblWidth.DXA);
        tcW.setW(BigInteger.valueOf(width));
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize, boolean bold,
                                         ParagraphAlignment alignment, int extraBreaks) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        String content = text == null ? "" : text;
        if (extraBreaks > 0) {
            StringBuilder builder = new StringBuilder(content);
            for (int i = 0; i < extraBreaks; i++) {
                builder.append("\n");
            }
            content = builder.toString();
        }
        setRunTextWithBreaks(run, content);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);
        run.setBold(bold);
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

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
    }
}
