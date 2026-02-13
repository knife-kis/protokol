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

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

final class ObservationProtocolWordExporter {

    private static final int FONT_SIZE = 9;

    private ObservationProtocolWordExporter() {
    }

    static void export(File target, String year, VlkWordExporter.PlanRow row, String controlDate) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {
            configureLandscapePage(document);
            createHeader(document);
            addTitle(document, year, row.responsible());
            addMainTable(document, row, controlDate);
            addSecondPage(document, controlDate);

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
        pageSize.setOrient(org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation.LANDSCAPE);

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
        setGrid(headerTable, 3200, 3200, 3200);

        setCellText(headerTable.getRow(0).getCell(0),
                "Испытательная лаборатория ООО «ЦИТ»", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(headerTable.getRow(0).getCell(1),
                "Протокол по результатам наблюдения\nФ3 РИ ИЛ 2-2023", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(headerTable.getRow(0).getCell(2),
                "Дата утверждения бланка формуляра: 01.01.2023г.", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(headerTable.getRow(1).getCell(2), "Редакция № 1", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setPageCounterCellText(headerTable.getRow(2).getCell(2));

        mergeCellsVertically(headerTable, 0, 0, 2);
        mergeCellsVertically(headerTable, 1, 0, 2);
        styleAllCellParagraphs(headerTable, FONT_SIZE);
    }

    private static void addTitle(XWPFDocument document, String year, String responsible) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(20);
        run.setBold(true);
        run.setText("Протокол по результатам наблюдений");
        run.addBreak();
        run.setText("за период: " + year + " г.");
        run.addBreak();
        run.setText("Ответственные лица: " + responsible);
        run.addBreak();
        run.setText("Отчет проверил и утвердил: " + oppositeResponsible(responsible));

        addEmptyLine(document);
    }

    private static void addMainTable(XWPFDocument document, VlkWordExporter.PlanRow row, String controlDate) {
        XWPFTable table = document.createTable(2, 7);
        table.setWidth("100%");
        setFixedLayout(table);
        setGrid(table, 650, 1300, 2800, 1300, 1300, 1300, 2200);

        XWPFTableRow headerRow = table.getRow(0);
        setCellText(headerRow.getCell(0), "№\nп/п", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(1), "Дата проведения контроля", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(2), "Шифр контролируемого метода (методики) измерений", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(3), "Контролируемый показатель", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(4), "Сотрудники ИЛ, демонстрирующие метод (методику) измерений", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(5), "Сотрудник ИЛ, осуществляющий контроль за работой сотрудников ИЛ", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(headerRow.getCell(6), "Результат контроля на каждом этапе реализации метода (методики) измерений", ParagraphAlignment.CENTER, true, FONT_SIZE);

        XWPFTableRow row1 = table.getRow(1);
        setCellText(row1.getCell(0), "1", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(row1.getCell(1), controlDate == null ? "" : controlDate, ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(row1.getCell(2), methodCipher(row.event()), ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(row1.getCell(3), controlledIndicator(row.event()), ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(row1.getCell(4), row.responsible(), ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(row1.getCell(5), oppositeResponsible(row.responsible()), ParagraphAlignment.LEFT, false, FONT_SIZE);
        setControlResultCell(row1.getCell(6), controlStages(row.event()), oppositeResponsible(row.responsible()), controlDate);

        styleAllCellParagraphs(table, FONT_SIZE);
    }

    private static void addSecondPage(XWPFDocument document, String controlDate) {
        XWPFParagraph pageBreak = document.createParagraph();
        pageBreak.createRun().addBreak(BreakType.PAGE);

        XWPFParagraph familiarTitle = document.createParagraph();
        familiarTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun familiarRun = familiarTitle.createRun();
        familiarRun.setFontFamily("Arial");
        familiarRun.setFontSize(14);
        familiarRun.setBold(true);
        familiarRun.setText("ЛИСТ ОЗНАКОМЛЕНИЯ");

        XWPFTable familiarTable = document.createTable(5, 5);
        familiarTable.setWidth("100%");
        setFixedLayout(familiarTable);
        setGrid(familiarTable, 1400, 2200, 2200, 1400, 2400);

        setCellText(familiarTable.getRow(0).getCell(0), "Дата\nознакомления", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(familiarTable.getRow(0).getCell(1), "Должность", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(familiarTable.getRow(0).getCell(2), "Ф.И.О.", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(familiarTable.getRow(0).getCell(3), "Подпись", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(familiarTable.getRow(0).getCell(4), "Примечание", ParagraphAlignment.CENTER, true, FONT_SIZE);

        setCellText(familiarTable.getRow(1).getCell(0), "1", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(familiarTable.getRow(1).getCell(1), "2", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(familiarTable.getRow(1).getCell(2), "3", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(familiarTable.getRow(1).getCell(3), "4", ParagraphAlignment.CENTER, false, FONT_SIZE);
        setCellText(familiarTable.getRow(1).getCell(4), "5", ParagraphAlignment.CENTER, false, FONT_SIZE);

        setCellText(familiarTable.getRow(2).getCell(1), "Заведующий лабораторией", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(familiarTable.getRow(2).getCell(2), "Тарновский Максим Олегович", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(familiarTable.getRow(2).getCell(0), controlDate == null ? "" : controlDate, ParagraphAlignment.LEFT, false, FONT_SIZE);

        setCellText(familiarTable.getRow(3).getCell(1), "Инженер", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(familiarTable.getRow(3).getCell(2), "Белов Дмитрий Андреевич", ParagraphAlignment.LEFT, false, FONT_SIZE);
        setCellText(familiarTable.getRow(3).getCell(0), controlDate == null ? "" : controlDate, ParagraphAlignment.LEFT, false, FONT_SIZE);

        styleAllCellParagraphs(familiarTable, FONT_SIZE);

        addEmptyLine(document);
        addEmptyLine(document);

        XWPFParagraph distributionTitle = document.createParagraph();
        distributionTitle.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun distributionRun = distributionTitle.createRun();
        distributionRun.setFontFamily("Arial");
        distributionRun.setFontSize(14);
        distributionRun.setBold(true);
        distributionRun.setText("ЛИСТ РАССЫЛКИ ДОКУМЕНТОВ");

        XWPFTable distributionTable = document.createTable(4, 4);
        distributionTable.setWidth("100%");
        setFixedLayout(distributionTable);
        setGrid(distributionTable, 1700, 2200, 2300, 3000);

        setCellText(distributionTable.getRow(0).getCell(0), "Номер учтенной копии", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(distributionTable.getRow(0).getCell(1), "Ф.И.О., должность", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(distributionTable.getRow(0).getCell(2), "Подпись о получении учтенной копии, дата", ParagraphAlignment.CENTER, true, FONT_SIZE);
        setCellText(distributionTable.getRow(0).getCell(3), "Отметка об изъятии у получателя учтенной копии, уничтожении учтенной копии: подпись уполномоченного работника ИЛ, дата", ParagraphAlignment.CENTER, true, FONT_SIZE);

        styleAllCellParagraphs(distributionTable, FONT_SIZE);
    }

    private static String methodCipher(String event) {
        String prefix = "Контроль точности результатов измерений по";
        if (event == null) {
            return "";
        }
        if (event.startsWith(prefix)) {
            return event.substring(prefix.length()).trim();
        }
        return event;
    }

    private static String oppositeResponsible(String responsible) {
        if (responsible == null) {
            return "";
        }
        if (responsible.contains("Белов")) {
            return "Тарновский М.О.";
        }
        if (responsible.contains("Тарновский")) {
            return "Белов Д.А.";
        }
        return "";
    }

    private static String controlledIndicator(String event) {
        if (event == null) {
            return "";
        }
        return switch (event) {
            case "Контроль точности результатов измерений по ГОСТ 30494-2011" ->
                    "Относительная влажность воздуха, температура поверхности отопительного прибора, температура воздуха, температура внутренней поверхности ограждения, результирующая температура, скорость движения воздуха.";
            case "Контроль точности результатов измерений по МИ Ш.13-2022" ->
                    "Эквивалентный уровень звука, дБА\nМаксимальный уровень звука, дБА";
            case "Контроль точности результатов измерений по МР 2.6.1.0361-24" ->
                    "Средневзвешенное значение мощности дозы гамма-\nизлучения\nПлотность потока радона с поверхности грунта\nМощность дозы гамма-излучения";
            case "Контроль точности результатов измерений по МУ 2.6.1.0333-23" ->
                    "Мощность дозы гамма-излучения\nЭРОА радона,\nЭРОА торона,\nсреднегодовое значение ЭРОА изотопов радона";
            case "Контроль точности результатов измерений по «Методика измерения плотности потока радона с поверхности земли и строительных конструкций (свидетельство об аттестации МВИ № 40090.6К816)»" ->
                    "Плотность потока радона с поверхности";
            case "Контроль точности результатов измерений по Руководство по эксплуатации Экофизика-110А ПКДУ.411000.001.02 РЭ п.7.1, п.7.2 с приложением МИ ПКФ 12-006 п.2, п.5" ->
                    "Уровень звукового давления в октавных (третьоктавных)\nполосах частот в диапазоне 31,5–16 000 Гц (25–20 000 Гц),\nдБ\nЭквивалентный уровень звука, дБА";
            case "Контроль точности результатов измерений по МИ СС.09-2021" ->
                    "освещеннось, КЕО, коэффициент пульсации минимальную освещенность, неравномеерность естественного освещения, равномерность освещенности, средняя горизонтальная освещенность на уровне земли, средняя освещенность";
            case "Контроль точности результатов измерений по ГОСТ 27296-2012, ГОСТ Р ИСО 3382-2-2013, ГОСТ Р 56769-2015, ГОСТ Р 56770-2015" -> "";
            default -> "";
        };
    }

    private static String[] controlStages(String event) {
        if (event == null) {
            return new String[0];
        }
        return switch (event) {
            case "Контроль точности результатов измерений по ГОСТ 30494-2011" -> new String[]{
                    "Применение оборудования в соответствии с Руководствами по эксплуатации",
                    "Определение периода года",
                    "Измерение температуры, влажности и скорости движения воздуха",
                    "Измерение температуры внутренней поверхности стен, световых проемов, отопительных приборов.",
                    "Измерение результирующей температуры"
            };
            case "Контроль точности результатов измерений по МИ Ш.13-2022" -> new String[]{
                    "Проверка работоспособности (калибровки) шумомера: выполнена на месте применения до и после серии измерений; отклонение показаний от уровня калибровочного сигнала не превышает 0,5 дБ",
                    "Проверено отсутствие факторов, искажающих результаты; обеспечены требуемые условия по методике (в т.ч. контроль климатических параметров)",
                    "Контроль выполнения ключевых операций метода: выбор контрольных точек и установка микрофона выполнены строго по методике",
                    "Измерения эквивалентного и максимального уровней выполнены в полном объёме по методике"
            };
            case "Контроль точности результатов измерений по МР 2.6.1.0361-24" -> new String[]{
                    "Проверены СИ и условия измерений",
                    "Выполнена поисковая гамма-съёмка участка по прямолинейным профилям с зигзагообразным движением детектора",
                    "Проведены измерения МАЭД в контрольных точках на высоте 1 м от поверхности",
                    "Проведен контроль ППР с поверхности грунта"
            };
            case "Контроль точности результатов измерений по МУ 2.6.1.0333-23" -> new String[]{
                    "Проверена пригодность СИ",
                    "Выполнена поисковая гамма-съёмка всех помещений вдоль стен и пола с зигзагообразным движением детектора",
                    "Проведены измерения МАЭД в помещениях",
                    "Проведены измерения ЭРОА изотопов радона/торона"
            };
            case "Контроль точности результатов измерений по «Методика измерения плотности потока радона с поверхности земли и строительных конструкций (свидетельство об аттестации МВИ № 40090.6К816)»" -> new String[]{
                    "Подготовка к отбору проб",
                    "Отбор проб.",
                    "Измерение плотности потока радона"
            };
            case "Контроль точности результатов измерений по Руководство по эксплуатации Экофизика-110А ПКДУ.411000.001.02 РЭ п.7.1, п.7.2 с приложением МИ ПКФ 12-006 п.2, п.5" -> new String[]{
                    "Подготовка прибора к работе",
                    "Проверка климатических и иных условий перед началом проведения измерений и их контроль во время проведения измерений",
                    "Проверка работоспособности шумомера",
                    "Измерение эквивалентного уровня звука и уровней звукового давления в октавных (третьоктавных)\nполосах частот"
            };
            case "Контроль точности результатов измерений по МИ СС.09-2021" -> new String[]{
                    "Подготовка объекта и условий измерений",
                    "Определение плоскости измерений и выбор/разметка контрольных точек",
                    "Измерение и расчет показателей",
                    "Измерение и расчет показателей"
            };
            case "Контроль точности результатов измерений по ГОСТ 27296-2012, ГОСТ Р ИСО 3382-2-2013, ГОСТ Р 56769-2015, ГОСТ Р 56770-2015" -> new String[]{
                    "Проведение калибровки",
                    "Подготовка объекта и условий измерений выполнен",
                    "Размещение источника шума и точек измерений (воздушный шум) выполнено по методике — расстояния/число точек соблюдены",
                    "Измерение времени реверберации",
                    "Расчет показателей звукоизоляции",
                    "Измерение ударного шума"
            };
            default -> new String[0];
        };
    }

    private static void setControlResultCell(XWPFTableCell cell,
                                             String[] stages,
                                             String controller,
                                             String controlDate) {
        cell.removeParagraph(0);
        if (stages == null || stages.length == 0) {
            addCellParagraph(cell, "", ParagraphAlignment.LEFT, false, FONT_SIZE);
            return;
        }
        for (String stage : stages) {
            addCellParagraph(cell, stage, ParagraphAlignment.LEFT, false, FONT_SIZE);
            addCellParagraph(cell, "__________/" + (controller == null ? "" : controller), ParagraphAlignment.RIGHT, false, FONT_SIZE);
            addCellParagraph(cell, controlDate == null ? "" : controlDate, ParagraphAlignment.RIGHT, false, FONT_SIZE);
        }
    }

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            CTVMerge merge = tcPr.isSetVMerge() ? tcPr.getVMerge() : tcPr.addNewVMerge();
            merge.setVal(rowIndex == fromRow ? STMerge.RESTART : STMerge.CONTINUE);
            if (rowIndex > fromRow) {
                setCellText(cell, "", ParagraphAlignment.LEFT, false, FONT_SIZE);
            }
        }
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

    private static void setCellText(XWPFTableCell cell,
                                    String text,
                                    ParagraphAlignment alignment,
                                    boolean bold,
                                    int size) {
        cell.removeParagraph(0);
        addCellParagraph(cell, text, alignment, bold, size);
    }

    private static void addCellParagraph(XWPFTableCell cell,
                                         String text,
                                         ParagraphAlignment alignment,
                                         boolean bold,
                                         int size) {
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingBefore(0);
        paragraph.setSpacingAfter(0);

        XWPFRun run = paragraph.createRun();
        run.setFontFamily("Arial");
        run.setFontSize(size);
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

    private static void styleAllCellParagraphs(XWPFTable table, int size) {
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph paragraph : cell.getParagraphs()) {
                    if (paragraph.getRuns() == null) {
                        continue;
                    }
                    for (XWPFRun run : paragraph.getRuns()) {
                        run.setFontFamily("Arial");
                        run.setFontSize(size);
                    }
                }
            }
        }
    }
}
