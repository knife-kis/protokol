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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;

final class EquipmentControlSheetExporter {
    private static final String CONTROL_SHEET_NAME = "лист контроля оборудования.docx";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int TABLE_WIDTH = 12560;

    private EquipmentControlSheetExporter() {
    }

    static void generate(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveControlSheetFile(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            applyControlSheetHeader(document);
            addControlSheetSection(
                    document,
                    "Камера-01",
                    "631",
                    "Измерения фона активированного угля",
                    new String[][]{
                            {"", "БДБ-2009", "", "Выбираем среднее значение", ""},
                            {"", "БДБ-2010", "", "Выбираем среднее значение", ""}
                    },
                    7
            );
            addPageBreak(document);
            addControlSheetSection(
                    document,
                    "Камера-01",
                    "631",
                    "измерения по контрольному источнику №647",
                    new String[][]{
                            {"", "БДБ-2009", "от 93 имп/с до 123 имп/с", "", ""},
                            {"", "БДБ-2010", "от 93 имп/с до 123 имп/с", "", ""}
                    },
                    7
            );

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveControlSheetFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), CONTROL_SHEET_NAME);
    }

    private static void applyControlSheetHeader(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
        XWPFParagraph headerParagraph = header.createParagraph();
        headerParagraph.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(headerParagraph);
        XWPFRun headerRun = headerParagraph.createRun();
        headerRun.setFontFamily(FONT_NAME);
        headerRun.setFontSize(FONT_SIZE);
        headerRun.setText("Лист контроля оборудования");
        headerRun.addBreak();
        headerRun.setText("Ф1 РИ ИЛ 2-2023 ");
    }

    private static void addControlSheetSection(
            XWPFDocument document,
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
        titleRun.setText("Лист контроля оборудования");
        titleRun.setFontFamily(FONT_NAME);
        titleRun.setFontSize(FONT_SIZE);
        titleRun.setBold(true);

        addSpacer(document);

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
        configureTableLayout(controlTable, new int[]{2512, 2512, 2512, 2512, 2512});
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
