package ru.citlab24.protokol.tabs.qms;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

final class TechnicalMaintenanceJournalWordExporter {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    private static final String MAINTENANCE_DESCRIPTION = """
            Ежедневное ТО включает:
            а) Внешний осмотр дозиметра-радиометра.
            б) Удаление пыли и грязи с наружных поверхностей.
            Еженедельное ТО, кроме операций ежедневного ТО, включает проверку работоспособности дозиметра-радиометра
            - Установите элемент питания (батарею или аккумулятор) в батарейный отсек.
            - Проверка работы детекторов. Включите дозиметр-радиометр, для чего нажмите кнопку 1 (Рис.5). Пульт автоматически перейдет в режим счета по каналу 1 (Таб.8), Счет по всем каналам измерения происходит следующим образом: на индикаторе появятся цифры «00.00» и символы, соответствующие каналу измерения (Таб.8), и начнется счет, сопровождающийся звуковыми сигналами, пропорциональными скорости счета. На индикаторе каждые 0,5 с будет появляться текущее среднее значение
            МЭД. По окончанию счета производится звуковой сигнал длительностью 1 с и результаты измерения в течении времени измерения индицируются на табло. Затем результат измерения обновляется и т.д. При превышении скорости счета 104 с-1 время измерения сокращается.
            Для выбора канала измерения используйте кнопку 6 («Канал»). Выбор канала измерения происходит при последовательных нажатиях кнопки 6 в следующем порядке: В случае если к пульту подключен выносной блок
            детектирования БДГ-01:
            канал 1 (встроенные детекторы СБМ-32, символ на индикаторе «mSv/h»)⇒
            канал 4 (блок детектирования БДГ-01, символы на индикаторе «mSv/h» и «g1»)⇒
            канал 2 (встроенный детектор СИ-34ГМ, символ на индикаторе «mSv/h»)⇒
            канал 1 и т.д.
            Просмотр накопленной эквивалентной дозы. Для просмотра накопленной эквивалентной дозы (далее дозы) в режиме измерения по любому каналу нажмите кнопку 3 («Доза»). На индикаторе появится значение накопленной эквивалентной дозы, ее размерность «mSv» и символы «g2», «D». Для выхода из режима просмотра дозы повторно нажмите кнопку 3 («Доза»).

            Ежемесячное ТО, кроме операций еженедельного ТО, включает проверку состояния корпуса дозиметра-радиометра (надежная фиксация
            переключателя, надежное крепление составных частей дозиметра-радиометра, сохранность герметизирующих прокладок) и выносных блоков детектирования (отсутствие повреждений и трещин).

            Ежегодное ТО, кроме операций ежемесячного ТО, включает поверку дозиметра-радиометра

            При техническом обслуживании дозиметра-радиометра следует
            соблюдать правила техники безопасности, изложенные в разделе 5 Паспорта

            При хранении дозиметра-радиометра в случае его переконсервации производится ТО, которое включает внешний осмотр и проверку работоспособности дозиметра-радиометра в соответствии с еженедельным ТО

            Перед использованием дозиметра-радиометра по назначению после его хранения более 1 года необходимо произвести ТО в объеме ежегодного ТО
            """;

    private TechnicalMaintenanceJournalWordExporter() {
    }

    static void export(File target, int year, int number) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            configureLandscapePage(document);
            createHeader(document);
            addTitlePage(document, year, number);
            addResponsiblePage(document);
            addMaintenancePages(document, year);
            addDistributionPage(document);
            try (FileOutputStream out = new FileOutputStream(target)) {
                document.write(out);
            }
        }
    }

    private static void configureLandscapePage(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                ? document.getDocument().getBody().getSectPr()
                : document.getDocument().getBody().addNewSectPr();
        CTPageSz pageSize = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pageSize.setW(BigInteger.valueOf(16838));
        pageSize.setH(BigInteger.valueOf(11906));
        pageSize.setOrient(STPageOrientation.LANDSCAPE);
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
        setGrid(headerTable, 3500, 5500, 3500);

        setCellText(headerTable.getRow(0).getCell(0), "Испытательная лаборатория ООО «ЦИТ»", ParagraphAlignment.CENTER, 12, false);
        setCellText(headerTable.getRow(0).getCell(1), "Журнал результатов методом повторных исследований (испытаний) и измерений Ф2 РИ ИЛ 2-2023", ParagraphAlignment.CENTER, 12, false);
        setCellText(headerTable.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.", ParagraphAlignment.LEFT, 12, false);
        setCellText(headerTable.getRow(1).getCell(2), "Редакция № 1", ParagraphAlignment.LEFT, 12, false);
        setPageCounterCellText(headerTable.getRow(2).getCell(2));

        mergeCellsVertically(headerTable, 0, 0, 2);
        mergeCellsVertically(headerTable, 1, 0, 2);
    }

    private static void addTitlePage(XWPFDocument document, int year, int number) {
        for (int i = 0; i < 5; i++) {
            document.createParagraph().createRun().addBreak();
        }
        XWPFTable titleTable = document.createTable(1, 1);
        titleTable.setWidth("65%");
        titleTable.setTableAlignment(TableRowAlign.CENTER);
        setCellText(titleTable.getRow(0).getCell(0),
                "ЖУРНАЛ\nтехнического обслуживания\n\n№" + number + "/" + year,
                ParagraphAlignment.CENTER,
                28,
                true);

        for (int i = 0; i < 4; i++) {
            document.createParagraph().createRun().addBreak();
        }

        XWPFParagraph footer = document.createParagraph();
        footer.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = footer.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(14);
        run.setText("Начат: _______________________");
        run.addBreak();
        run.setText("Окончен: ____________________");
        run.addBreak();
        run.setText("Срок хранения в архиве: _______");

        document.createParagraph().createRun().addBreak(BreakType.PAGE);
    }

    private static void addResponsiblePage(XWPFDocument document) {
        XWPFTable table = document.createTable(12, 2);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 6500, 6500);

        setCellText(table.getRow(0).getCell(0), "Ответственный за ведение журнала,\nДолжность, Фамилия, Инициалы.", ParagraphAlignment.LEFT, 12, true);
        setCellText(table.getRow(0).getCell(1), "Образец подписи", ParagraphAlignment.LEFT, 12, true);
        setCellText(table.getRow(1).getCell(0), "Заведующий лабораторией, Тарновский М.О.", ParagraphAlignment.LEFT, 12, false);
        setCellText(table.getRow(2).getCell(0), "Допущены к ведению журнала\nДолжность, Фамилия, Инициалы.", ParagraphAlignment.LEFT, 12, true);
        setCellText(table.getRow(2).getCell(1), "Образец подписи", ParagraphAlignment.LEFT, 12, true);
        setCellText(table.getRow(3).getCell(0), "Заведующий лабораторией, Тарновский М.О.", ParagraphAlignment.LEFT, 12, false);
        setCellText(table.getRow(4).getCell(0), "Инженер Белов Д.А.", ParagraphAlignment.LEFT, 12, false);
        setCellText(table.getRow(6).getCell(0), "Список сокращений (при необходимости)", ParagraphAlignment.CENTER, 12, true);
        mergeCellsHorizontally(table, 6, 0, 1);
        setCellText(table.getRow(7).getCell(0), "Условное обозначение", ParagraphAlignment.CENTER, 12, true);
        setCellText(table.getRow(7).getCell(1), "Расшифровка условного обозначения", ParagraphAlignment.CENTER, 12, true);
        setCellText(table.getRow(8).getCell(0), "Удовлетворительно", ParagraphAlignment.LEFT, 12, false);
        setCellText(table.getRow(8).getCell(1), "удов", ParagraphAlignment.LEFT, 12, false);

        XWPFParagraph caption = document.createParagraph();
        caption.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = caption.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(16);
        run.setText("Регистрация результатов проверок ведения журнала");

        XWPFTable checks = document.createTable(4, 6);
        checks.setWidth("100%");
        setFixedLayout(checks);
        setGrid(checks, 1300, 2200, 1800, 2300, 2300, 2300);
        setCellText(checks.getRow(0).getCell(0), "Дата проверки", ParagraphAlignment.CENTER, 12, false);
        setCellText(checks.getRow(0).getCell(1), "Период проверенных записей", ParagraphAlignment.CENTER, 12, false);
        setCellText(checks.getRow(0).getCell(2), "Результат проверки", ParagraphAlignment.CENTER, 12, false);
        setCellText(checks.getRow(0).getCell(3), "Ф.И.О. проверяющего, подпись", ParagraphAlignment.CENTER, 12, false);
        setCellText(checks.getRow(0).getCell(4), "Ознакомление ответственного", ParagraphAlignment.CENTER, 12, false);
        setCellText(checks.getRow(0).getCell(5), "Отметка об устранении", ParagraphAlignment.CENTER, 12, false);

        document.createParagraph().createRun().addBreak(BreakType.PAGE);
    }

    private static void addMaintenancePages(XWPFDocument document, int year) {
        XWPFTable equipment = document.createTable(5, 2);
        equipment.setWidth("100%");
        setFixedLayout(equipment);
        setGrid(equipment, 3600, 9400);
        setCellText(equipment.getRow(0).getCell(0), "Наименование оборудования:", ParagraphAlignment.LEFT, 10, false);
        setCellText(equipment.getRow(0).getCell(1), "Дозиметр-радиометр ДРБП-03", ParagraphAlignment.LEFT, 10, false);
        setCellText(equipment.getRow(1).getCell(0), "Заводской №:", ParagraphAlignment.LEFT, 10, false);
        setCellText(equipment.getRow(1).getCell(1), "21115", ParagraphAlignment.LEFT, 10, false);
        setCellText(equipment.getRow(2).getCell(0), "Описание технического обслуживания:", ParagraphAlignment.LEFT, 10, true);
        mergeCellsHorizontally(equipment, 2, 0, 1);
        setCellText(equipment.getRow(3).getCell(0), MAINTENANCE_DESCRIPTION, ParagraphAlignment.LEFT, 10, false);
        mergeCellsHorizontally(equipment, 3, 0, 1);
        mergeCellsHorizontally(equipment, 4, 0, 1);

        int daysInYear = LocalDate.of(year, 1, 1).lengthOfYear();
        XWPFTable journal = document.createTable(daysInYear + 1, 5);
        journal.setWidth("100%");
        setFixedLayout(journal);
        setGrid(journal, 700, 1500, 4700, 2500, 3500);

        setCellText(journal.getRow(0).getCell(0), "№\nп/п", ParagraphAlignment.CENTER, 10, true);
        setCellText(journal.getRow(0).getCell(1), "Дата", ParagraphAlignment.CENTER, 10, true);
        setCellText(journal.getRow(0).getCell(2), "Отметка о проведении технического обслуживания, заключение*", ParagraphAlignment.CENTER, 10, true);
        setCellText(journal.getRow(0).getCell(3), "Подпись ответственного", ParagraphAlignment.CENTER, 10, true);
        setCellText(journal.getRow(0).getCell(4), "Примечание", ParagraphAlignment.CENTER, 10, true);
        setRepeatTableHeader(journal.getRow(0));

        LocalDate date = LocalDate.of(year, 1, 1);
        for (int i = 1; i <= daysInYear; i++) {
            XWPFTableRow row = journal.getRow(i);
            row.setHeight(240);
            setCellText(row.getCell(0), String.valueOf(i), ParagraphAlignment.CENTER, 10, false);
            setCellText(row.getCell(1), date.format(DATE_FORMAT), ParagraphAlignment.CENTER, 10, false);
            setCellText(row.getCell(2), "", ParagraphAlignment.LEFT, 10, false);
            setCellText(row.getCell(3), "", ParagraphAlignment.LEFT, 10, false);
            setCellText(row.getCell(4), resolveMaintenanceType(date), ParagraphAlignment.CENTER, 10, false);
            date = date.plusDays(1);
        }

        XWPFParagraph note = document.createParagraph();
        XWPFRun noteRun = note.createRun();
        noteRun.setFontFamily("Arial");
        noteRun.setFontSize(10);
        noteRun.setText("* «+» - техническое обслуживание проведено, оборудование исправно.");
        noteRun.addBreak();
        noteRun.setText("  «-» - техническое обслуживание проведено, оборудование неисправно.");

        document.createParagraph().createRun().addBreak(BreakType.PAGE);
    }

    private static String resolveMaintenanceType(LocalDate date) {
        if (date.getDayOfMonth() == 1) {
            return "Ежемесячное";
        }
        if (date.getDayOfWeek() == DayOfWeek.MONDAY) {
            return "Еженедельное";
        }
        return "Ежедневное";
    }

    private static void addDistributionPage(XWPFDocument document) {
        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = title.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(14);
        run.setText("ЛИСТ РАССЫЛКИ ДОКУМЕНТОВ");

        XWPFTable table = document.createTable(4, 4);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 1600, 2600, 3800, 5000);
        setCellText(table.getRow(0).getCell(0), "Номер учтенной копии", ParagraphAlignment.CENTER, 12, false);
        setCellText(table.getRow(0).getCell(1), "Ф.И.О., должность", ParagraphAlignment.CENTER, 12, false);
        setCellText(table.getRow(0).getCell(2), "Подпись о получении учтенной копии, дата", ParagraphAlignment.CENTER, 12, false);
        setCellText(table.getRow(0).getCell(3), "Отметка об изъятии у получателя учтенной копии, уничтожении учтенной копии: подпись уполномоченного работника ИЛ, дата", ParagraphAlignment.CENTER, 12, false);
    }

    private static void setRepeatTableHeader(XWPFTableRow row) {
        row.setRepeatHeader(true);
    }

    private static void setPageCounterCellText(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);

        XWPFRun textRun = paragraph.createRun();
        textRun.setFontFamily("Arial");
        textRun.setFontSize(12);
        textRun.setText("Количество страниц: ");
        addField(paragraph, "PAGE");

        XWPFRun separatorRun = paragraph.createRun();
        separatorRun.setFontFamily("Arial");
        separatorRun.setFontSize(12);
        separatorRun.setText(" / ");
        addField(paragraph, "NUMPAGES");
    }

    private static void addField(XWPFParagraph paragraph, String fieldName) {
        XWPFRun beginRun = paragraph.createRun();
        beginRun.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        XWPFRun instrRun = paragraph.createRun();
        instrRun.getCTR().addNewInstrText().setStringValue(fieldName);
        XWPFRun separateRun = paragraph.createRun();
        separateRun.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        XWPFRun valueRun = paragraph.createRun();
        valueRun.setText("1");
        XWPFRun endRun = paragraph.createRun();
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

    private static void setCellText(XWPFTableCell cell, String text, ParagraphAlignment alignment, int fontSize, boolean bold) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(fontSize);
        run.setBold(bold);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            if (i > 0) {
                run.addBreak();
            }
            run.setText(lines[i]);
        }
    }

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(rowIndex == fromRow ? STMerge.RESTART : STMerge.CONTINUE);
            if (rowIndex > fromRow) {
                setCellText(cell, "", ParagraphAlignment.LEFT, 12, false);
            }
        }
    }

    private static void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            if (colIndex == fromCol) {
                tcPr.addNewHMerge().setVal(STMerge.RESTART);
            } else {
                tcPr.addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }
}
