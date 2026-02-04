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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public final class MeasurementPlanExporter {
    private static final String PLAN_SHEET_NAME = "план измерений.docx";
    private static final String FONT_NAME = "Arial";
    private static final int TITLE_FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 10;
    private static final int SMALL_FONT_SIZE = 8;
    private static final int MAP_APPLICATION_ROW_INDEX = 22;
    private static final String OBJECT_PREFIX = "4. Наименование объекта:";
    private static final String ADDRESS_PREFIX = "Адрес объекта";
    private static final String PERFORMER_PREFIX = "3. Измерения провел, подпись:";
    private static final String NORMATIVE_SECTION_TITLE = "Сведения о нормативных документах";
    private static final String NORMATIVE_HEADER_TITLE = "Измеряемый показатель";

    private MeasurementPlanExporter() {
    }

    static void generate(File sourceFile, File mapFile, String workDeadline) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveMeasurementPlanFile(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);
        String objectName = resolveObjectName(mapFile);
        String objectAddress = resolveObjectAddress(mapFile);
        String performer = resolveMeasurementPerformer(mapFile);
        boolean isTarnovsky = performer.contains("Тарновский");
        String performerRole = isTarnovsky ? "Заведующий лабораторией" : "Инженер";
        String responsibleText = buildResponsibleText(performerRole, performer);
        List<NormativeRow> normativeRows = resolveNormativeRows(sourceFile);
        if (normativeRows.isEmpty()) {
            normativeRows.add(new NormativeRow("", ""));
        }
        String protocolDate = resolveProtocolDateFromSource(sourceFile);
        String deadlineText = workDeadline == null ? "" : workDeadline.trim();
        boolean hasWorkDeadline = !deadlineText.isEmpty();
        String planCreatorRole = isTarnovsky
                ? "Заместитель заведующий лабораторией"
                : "Заведующий лабораторией";
        String planCreatorName = isTarnovsky ? "Гаврилова М.Е." : "Тарновский М.О.";
        String planCreatorDate = lastCharacters(applicationNumber, 10);
        String copyRecipientRole = planCreatorRole.equals("Заместитель заведующий лабораторией")
                ? "Заведующий лабораторией"
                : "Инженер";
        String copyRecipientName = copyRecipientRole.equals("Заведующий лабораторией")
                ? "Тарновский М.О."
                : "Белов Д.А.";

        try (XWPFDocument document = new XWPFDocument()) {
            setLandscapeOrientation(document);
            applyStandardHeader(document);

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(title);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("План измерений");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(TITLE_FONT_SIZE);
            titleRun.setBold(true);

            XWPFTable applicationTable = document.createTable(1, 2);
            setTableCellText(applicationTable.getRow(0).getCell(0), "Заявка №", TABLE_FONT_SIZE, true);
            setTableCellText(applicationTable.getRow(0).getCell(1), applicationNumber, TABLE_FONT_SIZE, false);

            XWPFParagraph spacer = document.createParagraph();
            setParagraphSpacing(spacer);

            int dataRowsCount = normativeRows.size();
            XWPFTable mainTable = document.createTable(1 + dataRowsCount, 6);
            configureMainTableLayout(mainTable);
            setTableCellText(mainTable.getRow(0).getCell(0), "Наименование объекта измерений", TABLE_FONT_SIZE, true);
            setTableCellText(mainTable.getRow(0).getCell(1), "Адрес объекта измерений", TABLE_FONT_SIZE, true);
            setTableCellText(mainTable.getRow(0).getCell(2),
                    "Ответственный от ИЛ (должность, фамилия, имя, отчество)",
                    TABLE_FONT_SIZE, true);
            setTableCellText(mainTable.getRow(0).getCell(3), "Показатель", TABLE_FONT_SIZE, true);
            setTableCellText(mainTable.getRow(0).getCell(4),
                    "Метод (методика) измерений (шифр)",
                    TABLE_FONT_SIZE, true);
            setTableCellText(mainTable.getRow(0).getCell(5),
                    "Срок выполнения работ в области лабораторной деятельности и оформления проекта отчета",
                    TABLE_FONT_SIZE, true);

            for (int index = 0; index < dataRowsCount; index++) {
                int rowIndex = index + 1;
                NormativeRow row = normativeRows.get(index);
                setTableCellText(mainTable.getRow(rowIndex).getCell(3), row.indicator, TABLE_FONT_SIZE, false);
                setTableCellText(mainTable.getRow(rowIndex).getCell(4), row.method, TABLE_FONT_SIZE, false);
                if (index == 0) {
                    setTableCellText(mainTable.getRow(rowIndex).getCell(0), objectName, TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(1), objectAddress, TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(2), responsibleText, TABLE_FONT_SIZE, false);
                    String deadlineValue = deadlineText;
                    setTableCellText(mainTable.getRow(rowIndex).getCell(5), deadlineValue, TABLE_FONT_SIZE, false);
                } else {
                    setTableCellText(mainTable.getRow(rowIndex).getCell(0), "", TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(1), "", TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(2), "", TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(5), "", TABLE_FONT_SIZE, false);
                }
            }

            if (dataRowsCount > 1) {
                mergeCellsVertically(mainTable, 0, 1, dataRowsCount);
                mergeCellsVertically(mainTable, 1, 1, dataRowsCount);
                mergeCellsVertically(mainTable, 2, 1, dataRowsCount);
                mergeCellsVertically(mainTable, 5, 1, dataRowsCount);
            }
            if (!hasWorkDeadline) {
                for (int rowIndex = 1; rowIndex <= dataRowsCount; rowIndex++) {
                    highlightCell(mainTable.getRow(rowIndex).getCell(5));
                }
            }

            addParagraphText(document, "План измерений сформировал:");

            XWPFTable creatorTable = document.createTable(2, 4);
            stretchTableToFullWidth(creatorTable);
            setTableCellText(creatorTable.getRow(0).getCell(0), planCreatorRole, TABLE_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(0).getCell(1), "", TABLE_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(0).getCell(2), planCreatorName, TABLE_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(0).getCell(3), planCreatorDate, TABLE_FONT_SIZE, false);

            setTableCellText(creatorTable.getRow(1).getCell(0), "должность", SMALL_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(1).getCell(1), "подпись", SMALL_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(1).getCell(2), "Ф.И.О", SMALL_FONT_SIZE, false);
            setTableCellText(creatorTable.getRow(1).getCell(3), "дата", SMALL_FONT_SIZE, false);

            addParagraphText(document, "Копию плана получил:");

            XWPFTable recipientTable = document.createTable(2, 4);
            stretchTableToFullWidth(recipientTable);
            setTableCellText(recipientTable.getRow(0).getCell(0), copyRecipientRole, TABLE_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(0).getCell(1), "", TABLE_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(0).getCell(2), copyRecipientName, TABLE_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(0).getCell(3), planCreatorDate, TABLE_FONT_SIZE, false);

            setTableCellText(recipientTable.getRow(1).getCell(0), "должность", SMALL_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(1).getCell(1), "подпись", SMALL_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(1).getCell(2), "Ф.И.О", SMALL_FONT_SIZE, false);
            setTableCellText(recipientTable.getRow(1).getCell(3), "дата", SMALL_FONT_SIZE, false);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    public static File resolveMeasurementPlanFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), PLAN_SHEET_NAME);
    }

    private static String buildResponsibleText(String role, String performer) {
        String trimmedRole = role == null ? "" : role.trim();
        String trimmedPerformer = performer == null ? "" : performer.trim();
        if (trimmedPerformer.isEmpty()) {
            return trimmedRole;
        }
        if (trimmedRole.isEmpty()) {
            return trimmedPerformer;
        }
        return trimmedRole + " " + trimmedPerformer;
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
            BigInteger width = (BigInteger) pageSize.getW();
            pageSize.setW(pageSize.getH());
            pageSize.setH(width);
        } else {
            pageSize.setW(BigInteger.valueOf(16840));
            pageSize.setH(BigInteger.valueOf(11900));
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
        setHeaderCellText(table.getRow(0).getCell(1), "План измерений\nФ7 РИ ИЛ 2-2023\n");
        setHeaderCellText(table.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(table.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellPageCount(table.getRow(2).getCell(2));

        mergeCellsVertically(table, 0, 0, 2);
        mergeCellsVertically(table, 1, 0, 2);
    }

    private static void configureMainTableLayout(XWPFTable table) {
        int[] columnWidths = {1600, 2000, 2000, 2000, 3360, 1600};

        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(12560));

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

    private static void stretchTableToFullWidth(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();
        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.PCT);
        tblW.setW(BigInteger.valueOf(5000));
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

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthDxa));
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

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize, boolean bold) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
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

    private static void addParagraphText(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);
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

    private static void highlightCell(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTShd shading = tcPr.isSetShd() ? tcPr.getShd() : tcPr.addNewShd();
        shading.setFill("FFFF00");
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
            String lastValue = "";
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.isEmpty()) {
                    if (!lastValue.isEmpty()) {
                        break;
                    }
                    continue;
                }
                lastValue = text;
            }
            return lastValue;
        } catch (Exception ignored) {
            return "";
        }
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

    private static String resolveObjectName(File mapFile) {
        String value = findValueByPrefix(mapFile, OBJECT_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "4. Наименование объекта");
        }
        return value;
    }

    private static String resolveObjectAddress(File mapFile) {
        String value = findValueByPrefix(mapFile, ADDRESS_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "Адрес объекта:");
        }
        return value;
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

    private static String trimLeadingPunctuation(String value) {
        int index = 0;
        while (index < value.length() && !Character.isLetterOrDigit(value.charAt(index))) {
            index++;
        }
        return value.substring(index).trim();
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

    private static List<NormativeRow> resolveNormativeRows(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new ArrayList<>();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return new ArrayList<>();
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int headerRowIndex = findNormativeHeaderRow(sheet, formatter);
            if (headerRowIndex < 0) {
                return new ArrayList<>();
            }
            List<NormativeRow> rows = new ArrayList<>();
            for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }
                String indicator = findFirstValueInRange(row, formatter, 0, 4);
                String method = findFirstValueInRange(row, formatter, 15, 25);
                if (indicator.contains(NORMATIVE_HEADER_TITLE)) {
                    continue;
                }
                if (indicator.isEmpty() && method.isEmpty()) {
                    break;
                }
                rows.add(new NormativeRow(indicator, method));
            }
            return rows;
        } catch (Exception ignored) {
            return new ArrayList<>();
        }
    }

    private static int findNormativeHeaderRow(Sheet sheet, DataFormatter formatter) {
        int sectionStart = -1;
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (sectionStart < 0 && text.contains(NORMATIVE_SECTION_TITLE)) {
                    sectionStart = row.getRowNum();
                }
                if (sectionStart >= 0 && text.contains(NORMATIVE_HEADER_TITLE)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static String findFirstValueInRange(Row row, DataFormatter formatter, int startCol, int endCol) {
        for (int col = startCol; col <= endCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty()) {
                return text;
            }
        }
        return "";
    }

    private static final class NormativeRow {
        private final String indicator;
        private final String method;

        private NormativeRow(String indicator, String method) {
            this.indicator = indicator == null ? "" : indicator;
            this.method = method == null ? "" : method;
        }
    }
}
