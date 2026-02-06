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
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Locale;

public final class RequestAnalysisSheetExporter {
    private static final String ANALYSIS_SHEET_NAME = "лист анализа заявки шумы.docx";
    private static final String FONT_NAME = "Arial";
    private static final int TITLE_FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 10;
    private static final int SMALL_FONT_SIZE = 8;
    private static final int PORTRAIT_TABLE_WIDTH_DXA = 9639;
    private static final String DATES_PREFIX = "2. Дата замеров:";
    private static final String PERFORMER_PREFIX = "3. Измерения провел, подпись:";
    private static final int MAP_APPLICATION_ROW_INDEX = 22;

    private RequestAnalysisSheetExporter() {
    }

    static void generate(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveAnalysisSheetFile(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);
        String performer = resolveMeasurementPerformer(mapFile);
        boolean isTarnovsky = performer.contains("Тарновский");
        String planCreatorRole = isTarnovsky
                ? "Заместитель заведующий лабораторией"
                : "Заведующий лабораторией";
        String responsibleRole = planCreatorRole.equals("Заместитель заведующий лабораторией")
                ? "Заведующий лабораторией"
                : "Инженер";
        String responsibleName = responsibleRole.equals("Заведующий лабораторией")
                ? "Тарновский М.О."
                : "Белов Д.А.";
        String responsibleDate = lastCharacters(applicationNumber, 10);
        List<String> measurementDates = resolveMeasurementDates(mapFile);
        String measurementDatesText = String.join(", ", measurementDates);

        try (XWPFDocument document = new XWPFDocument()) {
            applyStandardHeader(document);

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(title);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Лист анализа заявки");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(TITLE_FONT_SIZE);
            titleRun.setBold(true);

            addSpacer(document);

            XWPFTable applicationTable = document.createTable(1, 2);
            configureTableLayout(applicationTable, new int[]{4820, 4819});
            setTableCellText(applicationTable.getRow(0).getCell(0), "Заявка №",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.LEFT);
            setTableCellText(applicationTable.getRow(0).getCell(1), applicationNumber,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);

            addSpacer(document);

            XWPFTable analysisTable = document.createTable(6, 3);
            configureTableLayout(analysisTable, new int[]{6747, 1446, 1446});
            setTableCellText(analysisTable.getRow(0).getCell(0), "Анализ заявки, критерии",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(analysisTable.getRow(0).getCell(1), "ДА",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(analysisTable.getRow(0).getCell(2), "НЕТ",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

            String[] criteria = {
                    "Требования Заказчика надлежащим образом определены, документированы и правильно понимаются",
                    "ИЛ располагает возможностями и ресурсами для выполнения требований Заказчика",
                    "Выбраны соответствующие методы или методики, и они способны удовлетворить требования Заказчика",
                    "Дано согласие на передачу данных (в части отчетов о результатах) во ФГИС Росаккредитация",
                    "Угроза беспристрастности"
            };
            for (int index = 0; index < criteria.length; index++) {
                int rowIndex = index + 1;
                setTableCellText(analysisTable.getRow(rowIndex).getCell(0), criteria[index],
                        TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                setTableCellText(analysisTable.getRow(rowIndex).getCell(1), "",
                        TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
                setTableCellText(analysisTable.getRow(rowIndex).getCell(2), "",
                        TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }

            addSpacer(document);

            XWPFTable decisionTable = document.createTable(2, 4);
            configureTableLayout(decisionTable, new int[]{2892, 2249, 2249, 2249});
            setTableCellText(decisionTable.getRow(0).getCell(0), "Решение по заявке:",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(decisionTable.getRow(1).getCell(0), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(decisionTable.getRow(0).getCell(1), "Принять в работу",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(decisionTable.getRow(0).getCell(2), "Оформить отказ",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(decisionTable.getRow(0).getCell(3), "Запросить дополнительные материалы",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(decisionTable.getRow(1).getCell(1), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(decisionTable.getRow(1).getCell(2), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(decisionTable.getRow(1).getCell(3), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            mergeCellsVertically(decisionTable, 0, 0, 1);

            addSpacer(document);

            XWPFTable contractTable = document.createTable(1, 2);
            configureTableLayout(contractTable, new int[]{6747, 2892});
            setTableCellText(contractTable.getRow(0).getCell(0), "Договор с Заказчиком заключен:",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(contractTable.getRow(0).getCell(1), "ДА/НЕТ/НЕ ТРЕБУЕТСЯ",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            addSpacer(document);

            addParagraphText(document, "Лицо, ответственное за проведение анализа:");

            XWPFTable responsibleTable = document.createTable(2, 4);
            stretchTableToFullWidth(responsibleTable);
            setTableCellText(responsibleTable.getRow(0).getCell(0), responsibleRole,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(0).getCell(1), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(0).getCell(2), responsibleName,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(0).getCell(3), responsibleDate,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            setTableCellText(responsibleTable.getRow(1).getCell(0), "должность",
                    SMALL_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(1).getCell(1), "подпись",
                    SMALL_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(1).getCell(2), "Ф.И.О",
                    SMALL_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(responsibleTable.getRow(1).getCell(3), "дата",
                    SMALL_FONT_SIZE, false, ParagraphAlignment.CENTER);

            addSpacer(document);

            addParagraphText(document, "Обсуждения с Заказчиком:");

            XWPFTable discussionTable = document.createTable(4, 4);
            configureTableLayout(discussionTable, new int[]{2410, 2410, 2410, 2409});
            setTableCellText(discussionTable.getRow(0).getCell(0), "Дата переговоров (обсуждений)",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(discussionTable.getRow(0).getCell(1), "Содержание переговоров (обсуждений)",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(discussionTable.getRow(0).getCell(2), "Принятые в ИЛ решения",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(discussionTable.getRow(0).getCell(3),
                    "Подпись лица, ответственного за ведение обсуждений",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

            String datePrefix = measurementDatesText.isBlank() ? "" : " " + measurementDatesText;
            setTableCellText(discussionTable.getRow(1).getCell(0), responsibleDate,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(discussionTable.getRow(1).getCell(1),
                    "Согласование выезда с заказчиком на" + datePrefix,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(discussionTable.getRow(1).getCell(2),
                    "Запланировать выезд на" + datePrefix,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(discussionTable.getRow(1).getCell(3), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            setTableCellText(discussionTable.getRow(2).getCell(0), responsibleDate,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(discussionTable.getRow(2).getCell(1),
                    "Представитель заказчика укажет точки измерений на объекте в соответствии с заявкой ",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(discussionTable.getRow(2).getCell(2),
                    "Проводить работы только в присутствии с представителем заказчика",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(discussionTable.getRow(2).getCell(3), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            XWPFTableRow spacerRow = discussionTable.getRow(3);
            for (int colIndex = 0; colIndex < 4; colIndex++) {
                setTableCellText(spacerRow.getCell(colIndex), "", TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            }
            spacerRow.setHeight(4000);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    public static File resolveAnalysisSheetFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), ANALYSIS_SHEET_NAME);
    }

    private static void addParagraphText(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);
    }

    private static void addSpacer(XWPFDocument document) {
        XWPFParagraph spacer = document.createParagraph();
        setParagraphSpacing(spacer);
    }

    private static void stretchTableToFullWidth(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();
        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.PCT);
        tblW.setW(BigInteger.valueOf(5000));
    }

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(PORTRAIT_TABLE_WIDTH_DXA));

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

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthDxa));
    }

    private static void setTableCellText(XWPFTableCell cell,
                                         String text,
                                         int fontSize,
                                         boolean bold,
                                         ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);
        run.setBold(bold);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
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

    private static void applyStandardHeader(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        XWPFTable table = header.createTable(3, 3);
        configureHeaderTableLikeTemplate(table);

        setHeaderCellText(table.getRow(0).getCell(0), "Испытательная лаборатория ООО «ЦИТ»");
        setHeaderCellText(table.getRow(0).getCell(1), "Лист анализа заявки Ф69 ДП ИЛ 2-2023");
        setHeaderCellText(table.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(table.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellPageCount(table.getRow(2).getCell(2));

        mergeCellsVertically(table, 0, 0, 2);
        mergeCellsVertically(table, 1, 0, 2);
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

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(paragraph);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TITLE_FONT_SIZE);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
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
        run.setFontSize(TITLE_FONT_SIZE);

        appendField(paragraph, "PAGE");
        XWPFRun sep = paragraph.createRun();
        sep.setText(" / ");
        sep.setFontFamily(FONT_NAME);
        sep.setFontSize(TITLE_FONT_SIZE);

        appendField(paragraph, "NUMPAGES");
    }

    private static void appendField(XWPFParagraph paragraph, String instr) {
        XWPFRun runBegin = paragraph.createRun();
        runBegin.setFontFamily(FONT_NAME);
        runBegin.setFontSize(TITLE_FONT_SIZE);
        runBegin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun runInstr = paragraph.createRun();
        runInstr.setFontFamily(FONT_NAME);
        runInstr.setFontSize(TITLE_FONT_SIZE);
        var ctText = runInstr.getCTR().addNewInstrText();
        ctText.setStringValue(instr);
        ctText.setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFRun runSep = paragraph.createRun();
        runSep.setFontFamily(FONT_NAME);
        runSep.setFontSize(TITLE_FONT_SIZE);
        runSep.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun runText = paragraph.createRun();
        runText.setFontFamily(FONT_NAME);
        runText.setFontSize(TITLE_FONT_SIZE);
        runText.setText("1");

        XWPFRun runEnd = paragraph.createRun();
        runEnd.setFontFamily(FONT_NAME);
        runEnd.setFontSize(TITLE_FONT_SIZE);
        runEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(12));
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

    private static String resolveMeasurementPerformer(File mapFile) {
        String value = findValueByPrefix(mapFile, PERFORMER_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "3. Измерения провел, подпись");
        }
        return value;
    }

    private static String findValueByPrefix(File mapFile, String prefix) {
        if (mapFile == null || !mapFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = formatter.formatCellValue(cell).trim();
                    if (text.startsWith(prefix)) {
                        String tail = text.substring(prefix.length()).trim();
                        if (!tail.isEmpty()) {
                            return trimLeadingPunctuation(tail);
                        }
                        Cell next = row.getCell(cell.getColumnIndex() + 1);
                        if (next != null) {
                            String nextText = formatter.formatCellValue(next).trim();
                            if (!nextText.isEmpty()) {
                                return trimLeadingPunctuation(nextText);
                            }
                        }
                    }
                }
            }
        } catch (Exception ignored) {
            return "";
        }
        return "";
    }

    private static String trimLeadingPunctuation(String text) {
        if (text == null) {
            return "";
        }
        int index = 0;
        while (index < text.length()) {
            char ch = text.charAt(index);
            if (Character.isLetterOrDigit(ch)) {
                break;
            }
            if (!Character.isWhitespace(ch)) {
                index++;
                continue;
            }
            index++;
        }
        return text.substring(index).trim();
    }

    private static String lastCharacters(String value, int count) {
        if (value == null) {
            return "";
        }
        String trimmed = value.trim();
        if (trimmed.length() <= count) {
            return trimmed;
        }
        return trimmed.substring(trimmed.length() - count);
    }

    private static List<String> resolveMeasurementDates(File mapFile) {
        String rawDates = findValueByPrefix(mapFile, DATES_PREFIX);
        if (rawDates.isBlank()) {
            rawDates = findValueByPrefix(mapFile, "2. Дата замеров");
        }
        if (rawDates.isBlank()) {
            return Collections.emptyList();
        }
        String[] parts = rawDates.split(",");
        List<String> dates = new ArrayList<>();
        for (String part : parts) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) {
                dates.add(trimmed);
            }
        }
        return dates;
    }
}
