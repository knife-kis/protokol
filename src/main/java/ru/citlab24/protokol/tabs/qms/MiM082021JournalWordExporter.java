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
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTextDirection;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTextDirection;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Locale;
import java.util.Random;

final class MiM082021JournalWordExporter {

    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    private MiM082021JournalWordExporter() {
    }

    static void export(File target, String year, int number, LocalDate controlDate) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            configureLandscapePage(document);
            createHeader(document);
            addTitlePage(document, year, number);
            addResponsiblePage(document);
            addMeasurementPage(document, controlDate);
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

    private static void addTitlePage(XWPFDocument document, String year, int number) {
        XWPFTable box = document.createTable(1, 1);
        box.setWidth("100%");
        XWPFTableCell cell = box.getRow(0).getCell(0);
        setCellText(cell,
                "ЖУРНАЛ\nрезультатов методом повторных исследований (испытаний) и измерений\n№" + number + "/" + year,
                ParagraphAlignment.CENTER,
                28,
                true);

        XWPFParagraph spacer = document.createParagraph();
        spacer.setSpacingBefore(0);
        spacer.setSpacingAfter(0);
        XWPFRun run = spacer.createRun();
        run.addBreak();
        run.addBreak();
        run.addBreak();
        run.addBreak();

        XWPFParagraph footerText = document.createParagraph();
        footerText.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun ft = footerText.createRun();
        ft.setFontFamily("Arial");
        ft.setFontSize(14);
        ft.setText("Начат: _______________________");
        ft.addBreak();
        ft.setText("Окончен: ____________________");
        ft.addBreak();
        ft.setText("Срок хранения в архиве: _______");

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
        XWPFRun cr = caption.createRun();
        cr.setFontFamily("Arial");
        cr.setFontSize(16);
        cr.setBold(true);
        cr.setText("Регистрация результатов проверок ведения журнала");

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

    private static void addMeasurementPage(XWPFDocument document, LocalDate controlDate) {
        Random random = new Random(controlDate.toEpochDay());
        double x1TempAir = randomBetween(random, 22.0, 24.0, 1);
        double x2TempAir = vary(random, x1TempAir, 0.1, 22.0);
        double x1TempSurface = randomBetween(random, 17.0, 19.0, 1);
        double x2TempSurface = vary(random, x1TempSurface, 0.1, 17.0);
        double x1Humidity = randomBetween(random, 35.0, 45.0, 1);
        double x2Humidity = vary(random, x1Humidity, randomFrom(random, 0.0, 0.1, 0.2, 0.3), 35.0);
        double x1AirSpeed = randomBetween(random, 0.10, 0.13, 2);
        double x2AirSpeed = Math.max(0.10, vary(random, x1AirSpeed, 0.01, 0.10));

        XWPFTable table = document.createTable(5, 12);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 500, 950, 1700, 900, 1200, 1200, 850, 850, 1200, 1300, 1000, 1800);
        setRowHeight(table.getRow(0), 2000);
        setRowHeight(table.getRow(1), 1900);
        setRowHeight(table.getRow(2), 1900);
        setRowHeight(table.getRow(3), 1900);
        setRowHeight(table.getRow(4), 1900);

        setCellText(table.getRow(0).getCell(0), "№ п/п", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(1), "Дата\nпроведения\nконтроля", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(2), "Контролируемый\nпоказатель,\nединицы измерения", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(3), "НД на методику", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(4), "Средство(а)\nизмерения", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(5), "Оператор(ы)\n(Фамилия И.О.)", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(6), "Результаты измерений\nТарновский", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(7), "Результаты измерений\nБелов", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(8), "Среднее значение Хср\n(Х1+Х2)/2", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(9), "Результат процедуры\n|Х1-Х2|", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(10), "Норматив\nконтроля", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(0).getCell(11), "Вывод, подпись ответственного", ParagraphAlignment.CENTER, 7, false);

        fillRow(table.getRow(1), "1", controlDate.format(DATE_FORMAT), "Температура воздуха, °C", format(x1TempAir), format(x2TempAir), avg(x1TempAir, x2TempAir), abs(x1TempAir - x2TempAir), "0,2");
        fillRow(table.getRow(2), "2", "", "Температура поверхности, °C", format(x1TempSurface), format(x2TempSurface), avg(x1TempSurface, x2TempSurface), abs(x1TempSurface - x2TempSurface), "0,179");
        fillRow(table.getRow(3), "3", "", "Относительная влажность воздуха, %", format(x1Humidity), format(x2Humidity), avg(x1Humidity, x2Humidity), abs(x1Humidity - x2Humidity), "3,5");
        fillRow(table.getRow(4), "4", "", "Скорость движения воздуха, м/с", format(x1AirSpeed), format(x2AirSpeed), avg(x1AirSpeed, x2AirSpeed), abs(x1AirSpeed - x2AirSpeed), "0,105");

        mergeCellsVertically(table, 1, 1, 4);
        mergeCellsVertically(table, 3, 1, 4);
        mergeCellsVertically(table, 4, 1, 4);
        mergeCellsVertically(table, 5, 1, 4);

        setCellText(table.getRow(1).getCell(3), "МИ М.08−2021", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(1).getCell(4), "Метеоскоп-М №569521\nМетеоскоп-М 708623", ParagraphAlignment.CENTER, 7, false);
        setCellText(table.getRow(1).getCell(5), "Тарновский М.О.\nБелов Д.А.", ParagraphAlignment.CENTER, 7, false);
        setVerticalText(table.getRow(1).getCell(3));
        setVerticalText(table.getRow(1).getCell(4));

        document.createParagraph().createRun().addBreak(BreakType.PAGE);
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

    private static void fillRow(XWPFTableRow row, String idx, String date, String indicator,
                                String x1, String x2, double avg, double diff, String norm) {
        setCellText(row.getCell(0), idx, ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(1), date, ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(2), indicator, ParagraphAlignment.LEFT, 7, false);
        setCellText(row.getCell(6), x1, ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(7), x2, ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(8), format(avg), ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(9), format(diff), ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(10), norm, ParagraphAlignment.CENTER, 7, false);
        setCellText(row.getCell(11), "удов", ParagraphAlignment.CENTER, 7, false);
    }

    private static void setVerticalText(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTextDirection td = tcPr.isSetTextDirection() ? tcPr.getTextDirection() : tcPr.addNewTextDirection();
        td.setVal(STTextDirection.BT_LR);
    }

    private static void setRowHeight(XWPFTableRow row, int height) {
        row.setHeight(height);
    }

    private static double randomBetween(Random random, double min, double max, int digits) {
        double span = max - min;
        double value = min + random.nextDouble() * span;
        double m = Math.pow(10, digits);
        return Math.round(value * m) / m;
    }

    private static double vary(Random random, double base, double delta, double min) {
        int op = random.nextInt(3) - 1;
        return Math.max(min, base + op * delta);
    }

    private static double randomFrom(Random random, double... values) {
        return values[random.nextInt(values.length)];
    }

    private static String format(double value) {
        String raw = String.format(Locale.US, "%.3f", value);
        raw = raw.replaceAll("0+$", "").replaceAll("\\.$", "");
        return raw.replace('.', ',');
    }

    private static double avg(double a, double b) {
        return (a + b) / 2.0;
    }

    private static double abs(double value) {
        return Math.abs(value);
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
