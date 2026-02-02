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
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.Collections;
import java.util.List;

final class SoundInsulationEquipmentIssuanceSheetExporter {
    private static final String ISSUANCE_SHEET_BASE_NAME = "лист выдачи приборов";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 9;
    private static final int SIGNATURE_FONT_SIZE = 8;
    private static final String PERFORMER_NAME = "Белов Д.А.";

    private SoundInsulationEquipmentIssuanceSheetExporter() {
    }

    static void generate(File protocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        SoundInsulationProtocolDataParser.ProtocolData data = SoundInsulationProtocolDataParser.parse(protocolFile);
        String measurementDate = resolvePrimaryMeasurementDate(data.measurementDates());
        String objectName = data.objectName();
        List<SoundInsulationProtocolDataParser.InstrumentEntry> instruments = data.instruments();

        File targetFile = resolveIssuanceSheetFile(mapFile);
        writeIssuanceSheet(targetFile, objectName, PERFORMER_NAME, instruments, measurementDate);
    }

    static List<File> resolveIssuanceSheetFiles(File mapFile, File protocolFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }
        return List.of(resolveIssuanceSheetFile(mapFile));
    }

    private static void writeIssuanceSheet(File targetFile,
                                           String objectName,
                                           String performer,
                                           List<SoundInsulationProtocolDataParser.InstrumentEntry> instruments,
                                           String measurementDate) {
        try (XWPFDocument document = new XWPFDocument()) {
            applyStandardHeader(document);

            addCenteredParagraph(document, "Место использования прибора: " + safe(objectName), true);
            addCenteredParagraph(document,
                    "Фамилия, инициалы лица, получившего приборы:  " + safe(performer),
                    true);
            addCenteredParagraph(document, "Лист выдачи приборов", true);

            int rows = 1 + instruments.size() + 2;
            XWPFTable table = document.createTable(rows, 5);
            configureIssuanceTableLayout(table);

            setTableCellText(table.getRow(0).getCell(0), "Наименование\nоборудования", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(1), "Зав. №", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(2),
                    "Получение прибора, отметка о проведении технического обслуживания и (или) проверки перед выдачей",
                    TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(3),
                    "Возврат оборудования, Отметка о проведении технического обслуживания и (или) проверки перед возвратом оборудования и после его транспортировки",
                    TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(4), "Примечание", TABLE_FONT_SIZE);

            int rowIndex = 1;
            for (SoundInsulationProtocolDataParser.InstrumentEntry instrument : instruments) {
                setTableCellText(table.getRow(rowIndex).getCell(0), safe(instrument.name()), TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(1), safe(instrument.serialNumber()), TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(2), "", TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(3), "", TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);
                rowIndex++;
            }

            setTableCellText(table.getRow(rowIndex).getCell(0), "Дата", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(1), "", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(2), safe(measurementDate), TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(3), "___.___.20____", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);
            rowIndex++;

            setTableCellText(table.getRow(rowIndex).getCell(0), "Подпись", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(1), "", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(2),
                    "__________\n(лица, получающего прибор)", SIGNATURE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(3),
                    "________\n(лица, ответственного за метрологическое обеспечение)",
                    SIGNATURE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);

            addParagraphWithBreaks(document,
                    "Плановое обслуживание оборудования включает в себя осмотр на предмет отсутствия внешних повреждений, включение оборудования. " +
                            "Точный объем обслуживания определяется в эксплуатационной документации на оборудование, фиксируется в Плане обслуживания.\n" +
                            "При выдаче оборудования проводится проверка оборудования путем его осмотра на предмет целостности, отсутствия видимых дефектов, " +
                            "проверка уровня заряда аккумуляторных батарей (если применимо), наличия сигналов о неисправности при включении оборудования " +
                            "(если это не является составной частью технического обслуживания).\n" +
                            "Условные обозначения:+ - прибор получен (возвращен), техническое обслуживание и (или) промежуточная проверка (что предусмотрено) " +
                            "проведены, результат – прибор пригоден для дальнейшей работы.");

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    private static void addCenteredParagraph(XWPFDocument document, String text, boolean bold) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
        run.setBold(bold);
    }

    private static void addParagraphWithBreaks(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
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
        setHeaderCellText(table.getRow(0).getCell(1), "Лист выдачи приборов Ф39 ДП ИЛ 2-2023");
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

    private static void configureIssuanceTableLayout(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(9639));

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
        int[] widths = {1648, 1038, 2381, 2381, 2191};
        for (int width : widths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < widths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, widths[colIndex]);
            }
        }
    }

    private static void setBorder(CTBorder b) {
        b.setVal(STBorder.SINGLE);
        b.setSz(BigInteger.valueOf(4));
        b.setSpace(BigInteger.ZERO);
        b.setColor("auto");
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
        CTTblWidth w = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        w.setType(STTblWidth.DXA);
        w.setW(BigInteger.valueOf(widthDxa));
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(p);

        XWPFRun r = p.createRun();
        r.setText(text != null ? text : "");
        r.setFontFamily(FONT_NAME);
        r.setFontSize(FONT_SIZE);
    }

    private static void setHeaderCellPageCount(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph p = cell.addParagraph();
        p.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(p);

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

        XWPFRun rText = p.createRun();
        rText.setFontFamily(FONT_NAME);
        rText.setFontSize(FONT_SIZE);
        rText.setText("1");

        XWPFRun rEnd = p.createRun();
        rEnd.setFontFamily(FONT_NAME);
        rEnd.setFontSize(FONT_SIZE);
        rEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
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

    private static File resolveIssuanceSheetFile(File mapFile) {
        return new File(mapFile.getParentFile(), ISSUANCE_SHEET_BASE_NAME + ".docx");
    }

    private static String resolvePrimaryMeasurementDate(List<String> measurementDates) {
        if (measurementDates == null || measurementDates.isEmpty()) {
            return "";
        }
        String value = measurementDates.get(0);
        return value == null ? "" : value;
    }

    private static String safe(String value) {
        return value == null ? "" : value;
    }
}
