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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
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
import java.util.ArrayList;
import java.util.List;

final class VlkWordExporter {

    private static final int FONT_SIZE = 10;

    private VlkWordExporter() {
    }

    static void export(File target, String year, List<PlanRow> additionalRows) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            configurePortraitPage(document);
            createHeader(document);
            addApprovalTable(document, year);
            addTitle(document, year);

            List<PlanRow> allRows = new ArrayList<>(mandatoryRows());
            allRows.addAll(additionalRows);
            addMainTable(document, allRows);

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
        setGrid(headerTable, 3000, 3000, 3000);

        setCellText(headerTable.getRow(0).getCell(0),
                "Испытательная лаборатория ООО «ЦИТ»",
                ParagraphAlignment.CENTER,
                true);

        setCellText(headerTable.getRow(0).getCell(1),
                "План мониторинга достоверности результатов лабораторной деятельности\nФ44 ДП ИЛ 2-2023",
                ParagraphAlignment.CENTER,
                true);

        setCellText(headerTable.getRow(0).getCell(2),
                "Дата утверждения бланка формуляра: 01.01.2023г",
                ParagraphAlignment.LEFT,
                false);
        setCellText(headerTable.getRow(1).getCell(2), "Редакция № 1", ParagraphAlignment.LEFT, false);
        setPageCounterCellText(headerTable.getRow(2).getCell(2));

        mergeCellsVertically(headerTable, 0, 0, 2);
        mergeCellsVertically(headerTable, 1, 0, 2);

        styleAllCellParagraphs(headerTable);
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

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(rowIndex == fromRow ? STMerge.RESTART : STMerge.CONTINUE);
            if (rowIndex > fromRow) {
                setCellText(cell, "", ParagraphAlignment.LEFT, false);
            }
        }
    }

    private static void addApprovalTable(XWPFDocument document, String year) {
        XWPFTable approvalTable = document.createTable(1, 3);
        approvalTable.setWidth("100%");
        setFixedLayout(approvalTable);
        setGrid(approvalTable, 3000, 3000, 3000);
        removeAllBorders(approvalTable);

        setCellText(approvalTable.getRow(0).getCell(2),
                "УТВЕРЖДАЮ:\nЗаведующий ИЛ\n__________________________\n(подпись, инициалы, фамилия)\n«11» января " + year + "г.",
                ParagraphAlignment.LEFT,
                false);

        styleAllCellParagraphs(approvalTable);
        addEmptyLine(document);
    }

    private static void addTitle(XWPFDocument document, String year) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(FONT_SIZE);
        run.setBold(true);
        run.setText("План мониторинга достоверности результатов лабораторной деятельности");
        run.addBreak();
        run.setText("на " + year + " г.");

        addEmptyLine(document);
    }

    private static void addMainTable(XWPFDocument document, List<PlanRow> rows) {
        XWPFTable table = document.createTable(rows.size() + 1, 6);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 650, 3300, 1700, 1500, 1700, 1800);

        XWPFTableRow headerRow = table.getRow(0);
        setCellText(headerRow.getCell(0), "№ п/п", ParagraphAlignment.CENTER, true);
        setCellText(headerRow.getCell(1), "Наименования мероприятий", ParagraphAlignment.CENTER, true);
        setCellText(headerRow.getCell(2), "Периодичность", ParagraphAlignment.CENTER, true);
        setCellText(headerRow.getCell(3), "Ответственный", ParagraphAlignment.CENTER, true);
        setCellText(headerRow.getCell(4), "Формы учета", ParagraphAlignment.CENTER, true);
        setCellText(headerRow.getCell(5), "Отметка о выполнении дата/подпись ответственного лица", ParagraphAlignment.CENTER, true);

        markRowAsRepeatableHeader(headerRow);

        for (int i = 0; i < rows.size(); i++) {
            PlanRow row = rows.get(i);
            XWPFTableRow tableRow = table.getRow(i + 1);
            setCellText(tableRow.getCell(0), String.valueOf(i + 1), ParagraphAlignment.CENTER, false);
            setCellText(tableRow.getCell(1), row.event(), ParagraphAlignment.LEFT, false);
            setCellText(tableRow.getCell(2), row.periodicity(), ParagraphAlignment.LEFT, false);
            setCellText(tableRow.getCell(3), row.responsible(), ParagraphAlignment.LEFT, false);
            setCellText(tableRow.getCell(4), row.accountingForm(), ParagraphAlignment.LEFT, false);
            setCellText(tableRow.getCell(5), row.executionMark(), ParagraphAlignment.LEFT, false);
        }

        styleAllCellParagraphs(table);
    }


    private static void markRowAsRepeatableHeader(XWPFTableRow row) {
        if (row.getCtRow().getTrPr() == null) {
            row.getCtRow().addNewTrPr();
        }
        if (row.getCtRow().getTrPr().sizeOfTblHeaderArray() == 0) {
            row.getCtRow().getTrPr().addNewTblHeader();
        }
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

    private static void removeAllBorders(XWPFTable table) {
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTTblBorders borders = tblPr.isSetTblBorders() ? tblPr.getTblBorders() : tblPr.addNewTblBorders();
        borders.addNewTop().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
        borders.addNewLeft().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
        borders.addNewBottom().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
        borders.addNewRight().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
        borders.addNewInsideH().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
        borders.addNewInsideV().setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder.NONE);
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

    private static void setCellText(XWPFTableCell cell, String text, ParagraphAlignment alignment, boolean bold) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
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

    private static void addEmptyLine(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(FONT_SIZE);
        run.setText(" ");
    }

    private static void styleAllCellParagraphs(XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    if (paragraph.getRuns() == null) {
                        continue;
                    }
                    for (XWPFRun run : paragraph.getRuns()) {
                        run.setFontFamily("Arial");
                        run.setFontSize(FONT_SIZE);
                    }
                }
            }
        }
    }

    static List<PlanRow> mandatoryRows() {
        List<PlanRow> rows = new ArrayList<>();
        rows.add(new PlanRow(
                "Контроль за соблюдением требований нормативной документации при проведении измерений",
                "При утверждении каждого результата или при проверке ведения записей Заведующим ИЛ или аудитором – 1 раз в год",
                "Тарновский Максим Олегович",
                "Протокол испытаний или акт аудита",
                ""
        ));
        rows.add(new PlanRow(
                "Контроль условий окружающей среды при проведении измерений",
                "Проводится проверка записей Заведующим ИЛ или аудитором – 1 раз в год",
                "Тарновский Максим Олегович",
                "Протокол испытаний или акт аудита",
                ""
        ));
        rows.add(new PlanRow(
                "Проверку(и) функционирования (техническое обслуживание) СИ, ИО, ВО",
                "Проверка записей Заведующим ИЛ или аудитором – 1 раз в год",
                "Тарновский Максим Олегович",
                "Акт аудита",
                ""
        ));
        rows.add(new PlanRow(
                "Условия хранения и сроки годности экземпляров КИ, ведение записей",
                "Проверка записей Заведующим ИЛ или аудитором – 1 раз в год",
                "Тарновский Максим Олегович",
                "Акт аудита",
                ""
        ));
        rows.add(new PlanRow(
                "Условия и сроки хранения материалов",
                "Проверка записей Заведующим ИЛ или аудитором – 1 раз в год",
                "Тарновский Максим Олегович",
                "Акт аудита",
                ""
        ));
        rows.add(new PlanRow(
                "Проверка работоспособности (проверка калибровки) шумомера-виброметра, анализатора спектра",
                "При каждой серии измерений, до и после ее проведения",
                "Тарновский М.О.\n\nБелов Дмитрий Андреевич",
                "Карта замеров",
                ""
        ));
        rows.add(new PlanRow(
                "Проверка работоспособности Дозиметра-радиометра ДРБП-03",
                "Ежедневно",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Журнал Технического обслуживания №1",
                ""
        ));
        rows.add(new PlanRow(
                "Проверка работоспособности весов электронных XE-600",
                "Ежедневно",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Журнал Технического обслуживания №2",
                ""
        ));
        rows.add(new PlanRow(
                "Проверка работы Альфа-радиометра РАА-20П2 с помощью контрольного альфа-источника типа с радионуклидом Ат-241",
                "Перед началом и после серии измерений (один объект контроля)",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Карта замеров",
                ""
        ));
        rows.add(new PlanRow(
                "Измерение фона Альфа-радиометра РАА-20П2",
                "Перед началом серии измерений, а также после измерения проб, активность которых различается более чем в 100 раз",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Карта замеров",
                ""
        ));
        rows.add(new PlanRow(
                "Проверка температуры регенерации в регенераторе активированного угля",
                "При каждом измерении",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Журнал регистрации результатов определения плотности потока радона",
                ""
        ));
        rows.add(new PlanRow(
                "Измерение фона комплекса измерительного для мониторинга радона «КАМЕРА-01»",
                "При установке комплекса на новое место и ежедневно, перед началом измерения активности радона в угле (после контроля)",
                "Тарновский М.О.\n\nБелов Д.А.",
                "Листе контроля оборудования",
                ""
        ));
        rows.add(new PlanRow(
                "Юстировка весов",
                "При выдаче ошибки: «потеряны данные юстировки»",
                "Тарновский М.О.",
                "Бланк реагирования",
                ""
        ));
        rows.add(new PlanRow(
                "Контроль повторяемости с использованием рабочих измерений по ГОСТ 27296-2012",
                "Не менее 25 раз в год",
                "Белов Д.А.",
                "Программа по построению контрольных карт Шухарта",
                ""
        ));
        return rows;
    }

    record PlanRow(String event, String periodicity, String responsible, String accountingForm, String executionMark) {
    }
}
