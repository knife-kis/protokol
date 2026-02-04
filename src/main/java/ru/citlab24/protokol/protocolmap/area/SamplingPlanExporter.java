package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
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
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class SamplingPlanExporter {
    private static final String PLAN_NAME = "план отбора.docx";
    private static final String FONT_NAME = "Arial";
    private static final int TITLE_FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 9;
    private static final int TITLE_PAGE_ROW_INDEX = 4;
    private static final int PPR_START_ROW_INDEX = 4;
    private static final int DEFAULT_POINT_COUNT = 1;

    private SamplingPlanExporter() {
    }

    static void generate(File mapFile) {
        // обратная совместимость (если раньше вызывали только с картой)
        generate(null, mapFile);
    }

    static void generate(File sourceProtocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveSamplingPlanFile(mapFile);
        String applicationNumber = RequestFormExporter.resolveApplicationNumberFromMap(mapFile);
        String registrationDate = resolveRegistrationDate(mapFile);
        String approvalDate = formatApprovalDate(registrationDate);
        ApprovalBlock approvalBlock = resolveApprovalBlock(sourceProtocolFile);

        // ВАЖНО: точки считаем из ИСХОДНОГО протокола (sheet "ППР"),
        // а если не получилось — пробуем из карты (на случай других сценариев)
        int pointCount = resolveSamplingPointCount(sourceProtocolFile, mapFile);

        PerformerInfo performerInfo = resolvePerformerInfo(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            applySamplingHeader(document);

            // Блок "УТВЕРЖДАЮ" — таблица 2 колонки без границ, правый столбец по центру.
            XWPFTable approvalTable = document.createTable(1, 2);
            configureApprovalTableLayout(approvalTable);
            removeTableBorders(approvalTable);

            setTableCellText(approvalTable.getRow(0).getCell(0), "", TITLE_FONT_SIZE, false,
                    ParagraphAlignment.LEFT);

            setTableCellText(approvalTable.getRow(0).getCell(1), buildApprovalText(approvalDate, approvalBlock),
                    TITLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            XWPFParagraph spacer = document.createParagraph();
            setParagraphSpacing(spacer);
            spacer.createRun().addBreak();

            XWPFParagraph planTitle = document.createParagraph();
            planTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(planTitle);
            XWPFRun planTitleRun = planTitle.createRun();
            planTitleRun.setFontFamily(FONT_NAME);
            planTitleRun.setFontSize(TITLE_FONT_SIZE);
            planTitleRun.setBold(true);
            planTitleRun.setText("ПЛАН ОТБОРА");
            planTitleRun.addBreak();
            planTitleRun.setText("ПО ЗАЯВКЕ № " + applicationNumber);

            XWPFTable planTable = document.createTable(pointCount + 1, 8);
            configurePlanTableLayout(planTable);
            setTableCellText(planTable.getRow(0).getCell(0), "№ п/п", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(1),
                    "Место отбора проб (проведения измерений)", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(2),
                    "Точки отбора проб (проведения измерений)", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(3), "Наименование объекта испытаний",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(4),
                    "Определяемая характеристика (показатель)", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(5),
                    "Периодичность проведения работ", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(6),
                    "Документы, устанавливающие правила и методы отбора проб, измерений",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(7), "Примечание", TABLE_FONT_SIZE, true,
                    ParagraphAlignment.CENTER);

            for (int index = 0; index < pointCount; index++) {
                int rowIndex = index + 1;
                setTableCellText(planTable.getRow(rowIndex).getCell(0), String.valueOf(index + 1),
                        TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
                setTableCellText(planTable.getRow(rowIndex).getCell(2), "т" + (index + 1),
                        TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
                for (int colIndex = 1; colIndex < 8; colIndex++) {
                    if (colIndex == 2) {
                        continue;
                    }
                    setTableCellText(planTable.getRow(rowIndex).getCell(colIndex), "",
                            TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
                }
            }

            if (pointCount > 1) {
                mergeCellsVertically(planTable, 1, 1, pointCount);
                mergeCellsVertically(planTable, 3, 1, pointCount);
                mergeCellsVertically(planTable, 4, 1, pointCount);
                mergeCellsVertically(planTable, 5, 1, pointCount);
                mergeCellsVertically(planTable, 6, 1, pointCount);
                mergeCellsVertically(planTable, 7, 1, pointCount);
            }

            XWPFParagraph pageBreak = document.createParagraph();
            pageBreak.createRun().addBreak(BreakType.PAGE);

            XWPFParagraph secondPageBreak = document.createParagraph();
            secondPageBreak.createRun().addBreak(BreakType.PAGE);

            XWPFParagraph acceptTitle = document.createParagraph();
            acceptTitle.setAlignment(ParagraphAlignment.LEFT);
            setParagraphSpacing(acceptTitle);
            XWPFRun acceptRun = acceptTitle.createRun();
            acceptRun.setText("Принять в работу специалисту:");
            acceptRun.setFontFamily(FONT_NAME);
            acceptRun.setFontSize(TITLE_FONT_SIZE);

            XWPFTable performerTable = document.createTable(2, 4);
            stretchTableToFullWidth(performerTable);
            setTableCellText(performerTable.getRow(0).getCell(0), "ФИО исполнителя",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(0).getCell(1), "Должность", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(0).getCell(2), "Подпись", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(0).getCell(3), "Дата", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);

            setTableCellText(performerTable.getRow(1).getCell(0), performerInfo.fullName,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(1).getCell(1), performerInfo.role,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(1).getCell(2), "", TABLE_FONT_SIZE, false,
                    ParagraphAlignment.CENTER);
            setTableCellText(performerTable.getRow(1).getCell(3), registrationDate,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

            XWPFParagraph observerSpacer = document.createParagraph();
            setParagraphSpacing(observerSpacer);
            observerSpacer.createRun().addBreak();

            XWPFParagraph observerTitle = document.createParagraph();
            observerTitle.setAlignment(ParagraphAlignment.LEFT);
            setParagraphSpacing(observerTitle);
            XWPFRun observerRun = observerTitle.createRun();
            observerRun.setText("Наблюдение и контроль за соблюдением процедуры отбора проб " +
                    "(проведения измерений), если необходимо:");
            observerRun.setFontFamily(FONT_NAME);
            observerRun.setFontSize(TITLE_FONT_SIZE);

            XWPFTable observerTable = document.createTable(2, 4);
            stretchTableToFullWidth(observerTable);
            setTableCellText(observerTable.getRow(0).getCell(0), "ФИО исполнителя",
                    TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(0).getCell(1), "Должность", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(0).getCell(2), "Подпись", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(0).getCell(3), "Дата", TABLE_FONT_SIZE,
                    true, ParagraphAlignment.CENTER);

            setTableCellText(observerTable.getRow(1).getCell(0), "-", TABLE_FONT_SIZE, false,
                    ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(1).getCell(1), "-", TABLE_FONT_SIZE, false,
                    ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(1).getCell(2), "-", TABLE_FONT_SIZE, false,
                    ParagraphAlignment.CENTER);
            setTableCellText(observerTable.getRow(1).getCell(3), "-", TABLE_FONT_SIZE, false,
                    ParagraphAlignment.CENTER);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveSamplingPlanFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), PLAN_NAME);
    }

    private static void applySamplingHeader(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }

        // АЛЬБОМНАЯ ОРИЕНТАЦИЯ A4
        CTPageSz pgSz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pgSz.setOrient(STPageOrientation.LANDSCAPE);
        pgSz.setW(BigInteger.valueOf(16838)); // A4 landscape width (twips)
        pgSz.setH(BigInteger.valueOf(11906)); // A4 landscape height (twips)

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        XWPFTable table = header.createTable(3, 3);
        configureHeaderTableLikeTemplate(table);

        setHeaderCellText(table.getRow(0).getCell(0), "Испытательная лаборатория ООО «ЦИТ»");
        setHeaderCellText(table.getRow(0).getCell(1), "План отбора\nФ9 РИ ИЛ 2-2023");
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

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            XWPFRun run = paragraph.createRun();
            run.setText(lines[i]);
            run.setFontFamily(FONT_NAME);
            run.setFontSize(TITLE_FONT_SIZE);
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
        fieldValue.setFontSize(TITLE_FONT_SIZE);

        XWPFRun fieldEnd = paragraph.createRun();
        fieldEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);

        XWPFRun runTotal = paragraph.createRun();
        runTotal.setText(" / ");
        runTotal.setFontFamily(FONT_NAME);
        runTotal.setFontSize(TITLE_FONT_SIZE);

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
        fieldValueTotal.setFontSize(TITLE_FONT_SIZE);

        XWPFRun fieldEndTotal = paragraph.createRun();
        fieldEndTotal.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void configurePlanTableLayout(XWPFTable table) {
        int[] columnWidths = {700, 2300, 1200, 2300, 2100, 1800, 2800, 1300};
        configureTableLayout(table, columnWidths);
    }

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
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

    private static void setTableCellText(XWPFTableCell cell,
                                         String text,
                                         int fontSize,
                                         boolean bold,
                                         ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);

        String value = text == null ? "" : text;
        String[] lines = value.split("\\n", -1);

        for (int i = 0; i < lines.length; i++) {
            XWPFRun run = paragraph.createRun();
            run.setFontFamily(FONT_NAME);
            run.setFontSize(fontSize);
            run.setBold(bold);
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
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

    private static void stretchTableToFullWidth(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();
        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(12560));

        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(STJcTable.CENTER);
    }

    private static void removeTableBorders(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();
        CTTblBorders borders = pr.isSetTblBorders() ? pr.getTblBorders() : pr.addNewTblBorders();
        removeBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        removeBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        removeBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        removeBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        removeBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        removeBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());
    }

    private static void removeBorder(CTBorder border) {
        border.setVal(STBorder.NONE);
        border.setSz(BigInteger.ZERO);
        border.setSpace(BigInteger.ZERO);
        border.setColor("auto");
    }

    private static String buildApprovalText(String approvalDate, ApprovalBlock approvalBlock) {
        String dateValue = approvalDate == null || approvalDate.isBlank()
                ? "«17» февраля 2025 г."
                : approvalDate;
        return "УТВЕРЖДАЮ:" +
                "\n" + approvalBlock.title() +
                "\n" + approvalBlock.signature() +
                "\n(подпись, инициалы, фамилия)" +
                "\n" + dateValue;
    }

    private static ApprovalBlock resolveApprovalBlock(File sourceProtocolFile) {
        ApprovalBlock defaultBlock = new ApprovalBlock(
                "Заместитель заведующего лабораторией",
                "_____________________Гаврилова М.Е.");
        if (sourceProtocolFile == null || !sourceProtocolFile.exists()) {
            return defaultBlock;
        }
        try (InputStream in = new FileInputStream(sourceProtocolFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = workbook.getSheet("ППР");
            if (sheet == null) {
                return defaultBlock;
            }
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet) {
                if (row == null) {
                    continue;
                }
                StringBuilder rowText = new StringBuilder();
                boolean hasProtocolPrepared = false;
                short lastCellNum = row.getLastCellNum();
                for (int cellIndex = 0; cellIndex < lastCellNum; cellIndex++) {
                    String value = formatter.formatCellValue(row.getCell(cellIndex), evaluator).trim();
                    if (value.isEmpty()) {
                        continue;
                    }
                    if (rowText.length() > 0) {
                        rowText.append(' ');
                    }
                    rowText.append(value);
                    if (value.contains("Протокол подготовил:")) {
                        hasProtocolPrepared = true;
                    }
                }
                if (!hasProtocolPrepared && rowText.toString().contains("Протокол подготовил:")) {
                    hasProtocolPrepared = true;
                }
                if (hasProtocolPrepared) {
                    String lowerRow = rowText.toString().toLowerCase(Locale.ROOT);
                    if (lowerRow.contains("белов")) {
                        return new ApprovalBlock(
                                "заведующий лабораторией",
                                "___________________Тарновский М.О.");
                    }
                    if (lowerRow.contains("тарновский")) {
                        return defaultBlock;
                    }
                    return defaultBlock;
                }
            }
        } catch (Exception ex) {
            return defaultBlock;
        }
        return defaultBlock;
    }

    private record ApprovalBlock(String title, String signature) {
    }

    private static String resolveRegistrationDate(File mapFile) {
        File registrationFile = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mapFile);
        if (registrationFile == null || !registrationFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(registrationFile);
             XWPFDocument document = new XWPFDocument(in)) {
            if (document.getTables().isEmpty()) {
                return "";
            }
            XWPFTable table = document.getTables().get(0);
            if (table.getNumberOfRows() <= 1) {
                return "";
            }
            XWPFTableCell cell = table.getRow(1).getCell(0);
            String text = cell == null ? "" : cell.getText();
            return extractDateValue(text);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String extractDateValue(String text) {
        if (text == null) {
            return "";
        }
        Pattern pattern = Pattern.compile("от\\s*(\\d{2}\\.\\d{2}\\.\\d{4})", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(text);
        if (matcher.find()) {
            return matcher.group(1);
        }
        Pattern fallback = Pattern.compile("(\\d{2}\\.\\d{2}\\.\\d{4})");
        Matcher fallbackMatcher = fallback.matcher(text);
        if (fallbackMatcher.find()) {
            return fallbackMatcher.group(1);
        }
        return "";
    }

    private static String formatApprovalDate(String dateValue) {
        if (dateValue == null || dateValue.isBlank()) {
            return "";
        }
        try {
            LocalDate date = LocalDate.parse(dateValue.trim(), DateTimeFormatter.ofPattern("dd.MM.yyyy"));
            String monthName = monthName(date.getMonthValue());
            return String.format("«%02d» %s %d г.", date.getDayOfMonth(), monthName, date.getYear());
        } catch (DateTimeParseException ex) {
            return dateValue;
        }
    }

    private static String monthName(int month) {
        return switch (month) {
            case 1 -> "января";
            case 2 -> "февраля";
            case 3 -> "марта";
            case 4 -> "апреля";
            case 5 -> "мая";
            case 6 -> "июня";
            case 7 -> "июля";
            case 8 -> "августа";
            case 9 -> "сентября";
            case 10 -> "октября";
            case 11 -> "ноября";
            case 12 -> "декабря";
            default -> "";
        };
    }

    private static int resolveSamplingPointCount(File sourceProtocolFile, File mapFile) {
        int fromProtocol = countSamplingPointsFromFile(sourceProtocolFile);
        if (fromProtocol > 0) {
            return fromProtocol;
        }
        int fromMap = countSamplingPointsFromFile(mapFile);
        if (fromMap > 0) {
            return fromMap;
        }
        return DEFAULT_POINT_COUNT;
    }

    private static int countSamplingPointsFromFile(File file) {
        if (file == null || !file.exists()) {
            return 0;
        }
        try (InputStream in = new FileInputStream(file);
             Workbook workbook = WorkbookFactory.create(in)) {

            Sheet sheet = findSheet(workbook, "ППР");
            if (sheet == null) {
                return 0;
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            // По ТЗ: стартуем с строки 5 (B5-J5), но делаем страховку: если там нет "Точка" — ищем первую строку с "Точка"
            int startRow = PPR_START_ROW_INDEX;
            String startText = readMergedCellValue(sheet, startRow, 1, formatter, evaluator).toLowerCase(Locale.ROOT);
            if (!startText.contains("точка")) {
                int last = sheet.getLastRowNum();
                for (int r = 0; r <= Math.min(last, 200); r++) {
                    String t = readMergedCellValue(sheet, r, 1, formatter, evaluator).toLowerCase(Locale.ROOT);
                    if (t.contains("точка")) {
                        startRow = r;
                        break;
                    }
                }
            }

            int lastRow = sheet.getLastRowNum();
            int count = 0;

            for (int rowIndex = startRow; rowIndex <= lastRow; rowIndex++) {
                String text = readMergedCellValue(sheet, rowIndex, 1, formatter, evaluator)
                        .toLowerCase(Locale.ROOT);

                if (text.contains("точка")) {
                    count++;
                    continue;
                }
                if (count > 0) {
                    break;
                }
            }
            return count;
        } catch (Exception ex) {
            return 0;
        }
    }

    private static Sheet findSheet(Workbook workbook, String name) {
        if (workbook == null || name == null) {
            return null;
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet.getSheetName().equalsIgnoreCase(name)) {
                return sheet;
            }
        }
        return null;
    }
    private static void configureApprovalTableLayout(XWPFTable table) {
        // Левый столбец широкий (пустой), правый — под блок "УТВЕРЖДАЮ"
        int[] widths = {7000, 5560}; // суммарно 12560
        configureTableLayout(table, widths);
    }

    private static String readMergedCellValue(Sheet sheet,
                                              int rowIndex,
                                              int colIndex,
                                              DataFormatter formatter,
                                              FormulaEvaluator evaluator) {
        if (sheet == null) {
            return "";
        }
        CellRangeAddress region = findMergedRegion(sheet, rowIndex, colIndex);
        if (region != null) {
            Row row = sheet.getRow(region.getFirstRow());
            if (row == null) {
                return "";
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(region.getFirstColumn());
            return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        org.apache.poi.ss.usermodel.Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String normalizeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }

    private static PerformerInfo resolvePerformerInfo(File mapFile) {
        PerformerInfo info = PerformerInfo.empty();
        if (mapFile == null || !mapFile.exists()) {
            return info;
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return info;
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(TITLE_PAGE_ROW_INDEX);
            if (row == null) {
                return info;
            }
            DataFormatter formatter = new DataFormatter();
            for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
                org.apache.poi.ss.usermodel.Cell cell = row.getCell(colIndex);
                String text = formatter.formatCellValue(cell);
                if (text.contains("Гаврилова")) {
                    return PerformerInfo.tarnovsky();
                }
                if (text.contains("Тарновский")) {
                    return PerformerInfo.belov();
                }
            }
        } catch (Exception ex) {
            return info;
        }
        return info;
    }

    private static final class PerformerInfo {
        private final String fullName;
        private final String role;

        private PerformerInfo(String fullName, String role) {
            this.fullName = fullName;
            this.role = role;
        }

        private static PerformerInfo tarnovsky() {
            return new PerformerInfo("Тарновский Максим Олегович", "Заведующий лабораторией");
        }

        private static PerformerInfo belov() {
            return new PerformerInfo("Белов Дмитрий Андреевич", "инженер");
        }

        private static PerformerInfo empty() {
            return new PerformerInfo("", "");
        }
    }
}
