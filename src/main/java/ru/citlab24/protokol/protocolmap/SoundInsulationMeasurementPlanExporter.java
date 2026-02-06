package ru.citlab24.protokol.protocolmap;

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

final class SoundInsulationMeasurementPlanExporter {
    private static final String PLAN_SHEET_NAME = "план измерений.docx";
    private static final String FONT_NAME = "Arial";
    private static final int TITLE_FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 10;
    private static final int SMALL_FONT_SIZE = 8;
    private static final String PERFORMER_NAME = "Белов Д.А.";

    private SoundInsulationMeasurementPlanExporter() {
    }

    static void generate(File protocolFile, File mapFile, String workDeadline) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveMeasurementPlanFile(mapFile);
        SoundInsulationProtocolDataParser.ProtocolData data = SoundInsulationProtocolDataParser.parse(protocolFile);
        String applicationNumber = data.applicationNumber();
        if (applicationNumber.isBlank()) {
            applicationNumber = data.registrationNumber();
        }
        String objectName = data.objectName();
        String objectAddress = data.objectAddress();
        String performer = PERFORMER_NAME;
        boolean isTarnovsky = performer.contains("Тарновский");
        String performerRole = isTarnovsky ? "Заведующий лабораторией" : "Инженер";
        String responsibleText = buildResponsibleText(performerRole, performer);
        String protocolDate = data.protocolDate();
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
        List<PlanRow> rows = List.of(new PlanRow("Звукоизоляция ограждающих конструкций",
                data.measurementMethods()));

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

            int dataRowsCount = rows.size();
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
                PlanRow row = rows.get(index);
                setTableCellText(mainTable.getRow(rowIndex).getCell(3), row.indicator, TABLE_FONT_SIZE, false);
                setTableCellText(mainTable.getRow(rowIndex).getCell(4), row.method, TABLE_FONT_SIZE, false);
                if (index == 0) {
                    setTableCellText(mainTable.getRow(rowIndex).getCell(0), objectName, TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(1), objectAddress, TABLE_FONT_SIZE, false);
                    setTableCellText(mainTable.getRow(rowIndex).getCell(2), responsibleText, TABLE_FONT_SIZE, false);
                    String deadlineValue = hasWorkDeadline ? deadlineText : protocolDate;
                    setTableCellText(mainTable.getRow(rowIndex).getCell(5), deadlineValue, TABLE_FONT_SIZE, false);
                }
            }

            mergeCellsVertically(mainTable, 0, 1, dataRowsCount);
            mergeCellsVertically(mainTable, 1, 1, dataRowsCount);
            mergeCellsVertically(mainTable, 2, 1, dataRowsCount);
            mergeCellsVertically(mainTable, 5, 1, dataRowsCount);

            XWPFParagraph spacer2 = document.createParagraph();
            setParagraphSpacing(spacer2);

            XWPFParagraph creatorCaption = document.createParagraph();
            setParagraphSpacing(creatorCaption);
            creatorCaption.createRun().setText("План измерений сформировал:");

            XWPFTable approvalTable = document.createTable(2, 4);
            configureApprovalTableLayout(approvalTable);
            setTableCellText(approvalTable.getRow(0).getCell(0), planCreatorRole,
                    TABLE_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(0).getCell(1), "",
                    TABLE_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(0).getCell(2), planCreatorName,
                    TABLE_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(0).getCell(3), planCreatorDate,
                    TABLE_FONT_SIZE, false);

            setTableCellText(approvalTable.getRow(1).getCell(0), "должность",
                    SMALL_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(1).getCell(1), "подпись",
                    SMALL_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(1).getCell(2), "Ф.И.О",
                    SMALL_FONT_SIZE, false);
            setTableCellText(approvalTable.getRow(1).getCell(3), "дата",
                    SMALL_FONT_SIZE, false);

            XWPFParagraph spacer3 = document.createParagraph();
            setParagraphSpacing(spacer3);

            XWPFParagraph copyCaption = document.createParagraph();
            setParagraphSpacing(copyCaption);
            copyCaption.createRun().setText("Копию плана получил:");

            XWPFTable copyTable = document.createTable(2, 4);
            configureApprovalTableLayout(copyTable);
            setTableCellText(copyTable.getRow(0).getCell(0), copyRecipientRole,
                    TABLE_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(0).getCell(1), "",
                    TABLE_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(0).getCell(2), copyRecipientName,
                    TABLE_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(0).getCell(3), "",
                    TABLE_FONT_SIZE, false);

            setTableCellText(copyTable.getRow(1).getCell(0), "должность",
                    SMALL_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(1).getCell(1), "подпись",
                    SMALL_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(1).getCell(2), "Ф.И.О",
                    SMALL_FONT_SIZE, false);
            setTableCellText(copyTable.getRow(1).getCell(3), "дата",
                    SMALL_FONT_SIZE, false);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveMeasurementPlanFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), PLAN_SHEET_NAME);
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

    private static void configureMainTableLayout(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(14360));

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

        int[] widths = {2150, 2323, 2419, 1748, 1985, 1735};
        for (int width : widths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < widths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, widths[colIndex]);
            }
        }
    }

    private static void configureApprovalTableLayout(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.PCT);
        tblW.setW(BigInteger.valueOf(5000));

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
        int[] widths = {2692, 2692, 2692, 2692};
        for (int width : widths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < widths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, widths[colIndex]);
            }
        }
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
        setHeaderCellText(table.getRow(0).getCell(1), "План измерений Ф20 ДП ИЛ 2-2023");
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

    private static String buildResponsibleText(String role, String performer) {
        if (role == null && performer == null) {
            return "";
        }
        String textRole = role == null ? "" : role;
        String textName = performer == null ? "" : performer;
        if (textRole.isEmpty()) {
            return textName;
        }
        if (textName.isEmpty()) {
            return textRole;
        }
        return textRole + " " + textName;
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

    private record PlanRow(String indicator, String method) {
    }
}
