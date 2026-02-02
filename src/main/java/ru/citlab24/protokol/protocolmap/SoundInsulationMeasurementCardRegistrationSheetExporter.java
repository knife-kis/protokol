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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

final class SoundInsulationMeasurementCardRegistrationSheetExporter {
    private static final String REGISTRATION_SHEET_NAME = "лист регистрации карт замеров.docx";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final String PERFORMER_NAME = "Белов Д.А.";

    private SoundInsulationMeasurementCardRegistrationSheetExporter() {
    }

    static void generate(File protocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveRegistrationSheetFile(mapFile);
        SoundInsulationProtocolDataParser.ProtocolData data = SoundInsulationProtocolDataParser.parse(protocolFile);
        String applicationNumber = data.applicationNumber();
        String protocolNumber = data.protocolNumber();
        String cardNumber = data.registrationNumber();
        if (applicationNumber.isBlank()) {
            applicationNumber = cardNumber;
        }

        try (XWPFDocument document = new XWPFDocument()) {
            setLandscapeOrientation(document);
            applyStandardHeader(document);

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("Лист регистрации карт замеров");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(FONT_SIZE);

            XWPFTable table = document.createTable(2, 5);
            setTableCellText(table.getRow(0).getCell(0), "Номер заявки");
            setTableCellText(table.getRow(0).getCell(1), "Номер карты");
            setTableCellText(table.getRow(0).getCell(2), "Номер протокола");
            setTableCellText(table.getRow(0).getCell(3), "Фамилия работника - измерителя");
            setTableCellText(table.getRow(0).getCell(4),
                    "Подпись Заведующего ИЛ (заместителя Заведующего ИЛ)");

            setTableCellText(table.getRow(1).getCell(0), applicationNumber);
            setTableCellText(table.getRow(1).getCell(1), cardNumber);
            setTableCellText(table.getRow(1).getCell(2), protocolNumber);
            setTableCellText(table.getRow(1).getCell(3), PERFORMER_NAME);
            setTableCellText(table.getRow(1).getCell(4), "");

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (IOException ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveRegistrationSheetFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), REGISTRATION_SHEET_NAME);
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

    private static void setTableCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
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
        setHeaderCellText(table.getRow(0).getCell(1), "Лист регистрации карт замеров Ф77\nДП ИЛ 2-2023");
        setHeaderCellText(table.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(table.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellText(table.getRow(2).getCell(2), "Количество страниц: 1 / 1");

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

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
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
        run.setFontSize(FONT_SIZE);
        run.setText(text != null ? text : "");
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
}
