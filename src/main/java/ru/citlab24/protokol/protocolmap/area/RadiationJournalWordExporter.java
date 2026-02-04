package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TableWidthType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
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

import java.io.File;
import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.Locale;

final class RadiationJournalWordExporter {
    private static final String JOURNAL_FILE_NAME = "Журнал.docx";
    private static final String FONT_NAME = "Arial";
    private static final int TEXT_FONT_SIZE = 10;
    private static final int TABLE_FONT_SIZE = 9;
    private static final int TITLE_FONT_SIZE = 28;
    private static final int START_SK13_NUMBER = 17;
    private static final String HEADER_APPROVAL_DATE = "Дата утверждения бланка формуляра; 01.01.2023г.";
    private static final String HEADER_REVISION = "Редакция №1";

    private RadiationJournalWordExporter() {
    }

    static void generate(File sourceFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveJournalFile(mapFile);
        RadiationJournalData.ProtocolData data = RadiationJournalData.resolveProtocolData(sourceFile);

        try (XWPFDocument document = new XWPFDocument()) {
            setLandscapeOrientation(document);
            applyStandardHeader(document);

            buildTitlePage(document);
            buildResponsibleSection(document, data);
            buildCheckSection(document);
            buildProtocolInfoSection(document, data);
            buildPreparationSection(document);
            buildEnvironmentSection(document);
            List<String> sk13Numbers = buildSamplingSection(document, data);
            buildSamplingFooter(document, data, sk13Numbers);
            buildDistributionSection(document);
            buildChangeLogSection(document);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание файла, если не удалось сформировать документ
        }
    }

    static File resolveJournalFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), JOURNAL_FILE_NAME);
    }

    private static void buildTitlePage(XWPFDocument document) {
        XWPFTable titleTable = document.createTable(1, 1);
        stretchTableToFullWidth(titleTable);
        XWPFTableCell cell = titleTable.getRow(0).getCell(0);
        setTableCellText(cell,
                "ЖУРНАЛ\nРегистрации результатов определения плотности потока радона\n№5/2023",
                TITLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        addSpacer(document, 3);
        addParagraphText(document,
                "Начат: _______________________\nОкончен: ____________________\nСрок хранения в архиве: 3 года.",
                TEXT_FONT_SIZE,
                ParagraphAlignment.LEFT,
                false);
    }

    private static void buildResponsibleSection(XWPFDocument document, RadiationJournalData.ProtocolData data) {
        XWPFTable table = document.createTable(13, 2);
        stretchTableToFullWidth(table);

        setTableCellText(table.getRow(0).getCell(0),
                "Ответственный за ведение журнала,\nДолжность, Фамилия, Инициалы.",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(1),
                "Образец подписи", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        String preparedLine = data.preparedLine();
        setTableCellText(table.getRow(1).getCell(0),
                preparedLine, TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(1).getCell(1),
                "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(table.getRow(2).getCell(0),
                "Допущены к ведению журнала\nДолжность, Фамилия, Инициалы.",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(2).getCell(1),
                "Образец подписи", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        setTableCellText(table.getRow(3).getCell(0),
                preparedLine, TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(3).getCell(1),
                "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

        String tarnoLine = "Заведующий лабораторией Тарновский М.О.";
        String row5Text = preparedLine.toLowerCase(Locale.ROOT).contains("тарновский") ? "" : tarnoLine;
        setTableCellText(table.getRow(4).getCell(0),
                row5Text, TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(4).getCell(1),
                "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);

        for (int row = 5; row <= 5; row++) {
            setTableCellText(table.getRow(row).getCell(0),
                    "", TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(table.getRow(row).getCell(1),
                    "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }

        mergeCellsHorizontally(table, 6, 0, 1);
        setTableCellText(table.getRow(6).getCell(0),
                "Список сокращений (при необходимости)",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        setTableCellText(table.getRow(7).getCell(0),
                "Условное обозначение", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(1),
                "Расшифровка условного обозначения", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        for (int row = 8; row < 13; row++) {
            setTableCellText(table.getRow(row).getCell(0),
                    "", TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
            setTableCellText(table.getRow(row).getCell(1),
                    "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
    }

    private static void buildCheckSection(XWPFDocument document) {
        addParagraphText(document,
                "Регистрация результатов проверок ведения журнала",
                TEXT_FONT_SIZE,
                ParagraphAlignment.CENTER,
                false);

        XWPFTable table = document.createTable(4, 6);
        stretchTableToFullWidth(table);
        String[] headers = new String[]{
                "Дата проверки",
                "Период проверенных записей",
                "Результат проверки",
                "Ф.И.О. проверяющего, подпись",
                "Ознакомление ответственного",
                "Отметка об устранении"
        };
        for (int col = 0; col < headers.length; col++) {
            setTableCellText(table.getRow(0).getCell(col),
                    headers[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }
        for (int row = 1; row < 4; row++) {
            for (int col = 0; col < headers.length; col++) {
                setTableCellText(table.getRow(row).getCell(col),
                        "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }
        }
    }

    private static void buildProtocolInfoSection(XWPFDocument document, RadiationJournalData.ProtocolData data) {
        XWPFTable table = document.createTable(1, 4);
        stretchTableToFullWidth(table);
        setTableCellText(table.getRow(0).getCell(0),
                "Период наблюдений:", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(1),
                data.observationPeriod(), TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(0).getCell(2),
                "Населенный пункт:", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(3),
                data.settlement(), TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);

        addParagraphText(document, "Методика: " + data.method(),
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document, "Наименование заказчика: " + data.customerName(),
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document, "Заявка заказчика: " + data.customerRequest(),
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document,
                "Присутствие представителя Заказчика при отборе: " + data.customerPresence(),
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document, "Подготовка к отбору пробы:",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
    }

    private static void buildPreparationSection(XWPFDocument document) {
        XWPFTable table = document.createTable(4, 6);
        stretchTableToFullWidth(table);
        String[] headers = new String[]{
                "Дата регенерации",
                "Продолжительность регенерации",
                "t активированного угля, °С",
                "Помещение угля в СК-13 (+/-)",
                "Уголь пригоден для отбора (+)/не пригоден для отбора (-)",
                "Подпись ответственного ха подготовку к отбору пробы"
        };
        for (int col = 0; col < headers.length; col++) {
            setTableCellText(table.getRow(0).getCell(col),
                    headers[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }
        for (int row = 1; row < 4; row++) {
            for (int col = 0; col < headers.length; col++) {
                setTableCellText(table.getRow(row).getCell(col),
                        "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }
        }

        addParagraphText(document,
                "Условные обозначения: + выполнено;  - не выполнено. ",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document,
                "Требования к условиям окружающей среды:",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
    }

    private static void buildEnvironmentSection(XWPFDocument document) {
        XWPFTable table = document.createTable(10, 8);
        stretchTableToFullWidth(table);

        mergeCellsVertically(table, 0, 0, 1);
        mergeCellsVertically(table, 1, 0, 1);
        mergeCellsHorizontally(table, 0, 2, 6);
        mergeCellsVertically(table, 7, 0, 1);

        setTableCellText(table.getRow(0).getCell(0), "№ п/п", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(1),
                "Шифр методики/ описание деятельности", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(2),
                "Рабочие условия", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(7),
                "Дополнительные сведения", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        setTableCellText(table.getRow(1).getCell(2), "Т, °С", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(3), "Р, мм рт. ст", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(4),
                "Относительная влажность, %", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(5), "Напряжение сети, В", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(6), "Частота тока, Гц", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        mergeCellsVertically(table, 0, 2, 5);
        mergeCellsVertically(table, 1, 2, 5);
        mergeCellsHorizontally(table, 2, 2, 6);
        mergeCellsHorizontally(table, 4, 2, 6);
        mergeCellsVertically(table, 7, 2, 5);

        setTableCellText(table.getRow(2).getCell(0), "1.", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(2).getCell(1),
                "Методика измерений плотности потока радона с поверхности земли и " +
                        "строительных конструкций (свидетельство № 40090.6К816 об аттестации МВИ) " +
                        "при отборе образцов",
                TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(2).getCell(2), "НОРМА", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(3).getCell(2), "от минус 15°С до +40 °С",
                TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(3).getCell(3), "-", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(3).getCell(4), "-", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(3).getCell(5), "-", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(3).getCell(6), "-", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(4).getCell(2), "Фактически", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        for (int col = 2; col <= 6; col++) {
            setTableCellText(table.getRow(5).getCell(col), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
        setTableCellText(table.getRow(2).getCell(7),
                "Контроль осуществляется в течение времени пассивного отбора образцов, " +
                        "для чего используется сигнализатор, который световым сигналом оповещает " +
                        "работника ИЛ о выходе за пределы температурного \n" +
                        "режима, которые влияют на погрешность измерения. ",
                TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);

        mergeCellsVertically(table, 0, 6, 9);
        mergeCellsVertically(table, 1, 6, 9);
        mergeCellsHorizontally(table, 6, 2, 6);
        mergeCellsHorizontally(table, 8, 2, 6);
        mergeCellsVertically(table, 7, 6, 9);

        setTableCellText(table.getRow(6).getCell(0), "2.", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(6).getCell(1),
                "Проведение измерений по Методике измерений плотности потока радона с " +
                        "поверхности земли и строительных конструкций (свидетельство № 40090.6К816 об аттестации МВИ)",
                TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(table.getRow(6).getCell(2), "НОРМА", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(2),
                "от плюс 10°С до плюс 35 °С", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(3),
                "от 630 мм рт. ст. до 802 мм рт. ст.", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(4),
                "до 95 % при температуре плюс 30 °С", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(5), "220+22 -33", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(7).getCell(6), "50±1", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(8).getCell(2), "Фактически", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        for (int col = 2; col <= 6; col++) {
            setTableCellText(table.getRow(9).getCell(col), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
        setTableCellText(table.getRow(6).getCell(7),
                "Руководство по эксплуатации  \nФМКТ.136132.134 РЭ \n" +
                        "Комплекс измерительный для мониторинга радона «КАМЕРА-01» ",
                TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
    }

    private static List<String> buildSamplingSection(XWPFDocument document, RadiationJournalData.ProtocolData data) {
        addParagraphText(document, "Сведения об отборе:",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);

        int pointCount = Math.max(RadiationJournalData.DEFAULT_POINT_COUNT, data.pointCount());
        XWPFTable table = document.createTable(2 + pointCount, 9);
        stretchTableToFullWidth(table);

        mergeCellsVertically(table, 0, 0, 1);
        mergeCellsVertically(table, 1, 0, 1);
        mergeCellsVertically(table, 2, 0, 1);
        mergeCellsVertically(table, 3, 0, 1);
        mergeCellsVertically(table, 4, 0, 1);
        mergeCellsHorizontally(table, 0, 5, 6);
        mergeCellsVertically(table, 7, 0, 1);
        mergeCellsVertically(table, 8, 0, 1);

        setTableCellText(table.getRow(0).getCell(0),
                "Местоположение пункта наблюдения", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(1),
                "Подготовка участка (очистка от дерна, рыхление земли, перед установкой НК-32 прошло более 1 ч): +/-",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(2),
                "Номер СК-13", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(3),
                "Количество совместно экспонированных НК-32", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(4),
                "Место установки НК-32", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(5),
                "Дата и время экспонирования (с точностью до 1 мин)", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(5),
                "Начала", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(6),
                "Окончания", TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(7),
                "Отклонения, дополнения или исключения от условий отбора",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(8),
                "Пригоден образец после транспортировки  для последующих измерений: ДА/НЕТ",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        if (pointCount > 0) {
            mergeCellsVertically(table, 0, 2, 1 + pointCount);
            setTableCellText(table.getRow(2).getCell(0),
                    data.observationLocation(), TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
        }

        int seq = START_SK13_NUMBER;
        List<String> sk13Numbers = new java.util.ArrayList<>();
        for (int i = 0; i < pointCount; i++) {
            int rowIndex = 2 + i;
            XWPFTableRow row = table.getRow(rowIndex);
            setTableCellText(row.getCell(1), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            String sk13Number = String.format("%04d", seq++);
            sk13Numbers.add(sk13Number);
            setTableCellText(row.getCell(2), sk13Number,
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(3), "1", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(4), "т" + (i + 1),
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(5), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(6), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(7), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            setTableCellText(row.getCell(8), "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
        return sk13Numbers;
    }

    private static void buildSamplingFooter(XWPFDocument document,
                                            RadiationJournalData.ProtocolData data,
                                            List<String> sk13Numbers) {
        addParagraphText(document,
                "Условные обозначения: + выполнено; - не выполнено.",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);
        addParagraphText(document, "Отбор произвел: ",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);

        XWPFTable takerTable = document.createTable(2, 4);
        stretchTableToFullWidth(takerTable);
        for (int col = 0; col < 4; col++) {
            setTableCellText(takerTable.getRow(0).getCell(col), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
        String[] takerHeaders = new String[]{"Должность", "подпись", "Ф.И.О", "дата"};
        for (int col = 0; col < 4; col++) {
            setTableCellText(takerTable.getRow(1).getCell(col),
                    takerHeaders[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }

        addParagraphText(document, "На измерения образец принял: ",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);

        XWPFTable acceptTable = document.createTable(2, 4);
        stretchTableToFullWidth(acceptTable);
        for (int col = 0; col < 4; col++) {
            setTableCellText(acceptTable.getRow(0).getCell(col), "",
                    TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
        for (int col = 0; col < 4; col++) {
            setTableCellText(acceptTable.getRow(1).getCell(col),
                    takerHeaders[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }

        addParagraphText(document, "Измерения:",
                TEXT_FONT_SIZE, ParagraphAlignment.LEFT, false);

        int pointCount = Math.max(RadiationJournalData.DEFAULT_POINT_COUNT, data.pointCount());
        XWPFTable measurementTable = document.createTable(1 + pointCount, 5);
        stretchTableToFullWidth(measurementTable);
        String[] measurementHeaders = new String[]{
                "Номер СК-13",
                "Дата и время измерения",
                "Значение плотности потока радона, мБк/м^2с",
                "Погрешность измерения, мБк/м^2с",
                "Подпись измерителя"
        };
        for (int col = 0; col < measurementHeaders.length; col++) {
            setTableCellText(measurementTable.getRow(0).getCell(col),
                    measurementHeaders[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }
        for (int row = 1; row <= pointCount; row++) {
            for (int col = 0; col < measurementHeaders.length; col++) {
                setTableCellText(measurementTable.getRow(row).getCell(col),
                        "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }
        }
        for (int row = 1; row <= pointCount; row++) {
            String sk13Number = row - 1 < sk13Numbers.size() ? sk13Numbers.get(row - 1) : "";
            setTableCellText(measurementTable.getRow(row).getCell(0),
                    sk13Number, TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
        }
    }

    private static void buildDistributionSection(XWPFDocument document) {
        addParagraphText(document,
                "ЛИСТ РАССЫЛКИ ДОКУМЕНТОВ",
                TEXT_FONT_SIZE,
                ParagraphAlignment.CENTER,
                false);

        XWPFTable table = document.createTable(4, 4);
        stretchTableToFullWidth(table);
        String[] headers = new String[]{
                "Номер учтенной копии",
                "Ф.И.О., должность",
                "Подпись о получении учтенной копии, дата",
                "Отметка об изъятии у получателя учтенной копии, уничтожении учтенной копии: " +
                        "подпись уполномоченного работника ИЛ, дата"
        };
        for (int col = 0; col < headers.length; col++) {
            setTableCellText(table.getRow(0).getCell(col),
                    headers[col], TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        }
        for (int row = 1; row < 4; row++) {
            for (int col = 0; col < headers.length; col++) {
                setTableCellText(table.getRow(row).getCell(col),
                        "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }
        }
    }

    private static void buildChangeLogSection(XWPFDocument document) {
        addParagraphText(document,
                "ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ",
                TEXT_FONT_SIZE,
                ParagraphAlignment.CENTER,
                false);

        XWPFTable table = document.createTable(6, 8);
        stretchTableToFullWidth(table);

        mergeCellsVertically(table, 0, 0, 1);
        mergeCellsHorizontally(table, 0, 1, 4);
        mergeCellsVertically(table, 5, 0, 1);
        mergeCellsVertically(table, 6, 0, 1);
        mergeCellsVertically(table, 7, 0, 1);

        setTableCellText(table.getRow(0).getCell(0), "№\nизменения",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(1), "Номер листа (страницы), пункта",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(5), "Дата \nутверждения\n",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(6),
                "Ф.И.О., должность лица, внесшего изменения",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(0).getCell(7), "Подпись МК",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        setTableCellText(table.getRow(1).getCell(1), "измененного",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(2), "замененного",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(3), "нового",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
        setTableCellText(table.getRow(1).getCell(4), "аннулированного",
                TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

        for (int row = 2; row < 6; row++) {
            for (int col = 0; col < 8; col++) {
                setTableCellText(table.getRow(row).getCell(col),
                        "", TABLE_FONT_SIZE, false, ParagraphAlignment.CENTER);
            }
        }
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

        setHeaderCellText(table.getRow(0).getCell(0), "Испытательная лаборатория\nООО «ЦИТ»",
                ParagraphAlignment.CENTER);
        setHeaderCellText(table.getRow(0).getCell(1),
                "Журнал регистрации результатов\nопределения плотности потока радона\nФ6 РИ ИЛ 2-2023");
        setHeaderCellText(table.getRow(0).getCell(2), HEADER_APPROVAL_DATE, ParagraphAlignment.RIGHT);
        setHeaderCellText(table.getRow(1).getCell(2), HEADER_REVISION, ParagraphAlignment.RIGHT);
        setHeaderCellPageCount(table.getRow(2).getCell(2), ParagraphAlignment.RIGHT);

        mergeCellsVertically(table, 0, 0, 2);
        mergeCellsVertically(table, 1, 0, 2);
        tightenHeaderRowHeights(table);
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

    private static void applyMinimalHeaderCellMargins(CTTblPr pr) {
        CTTblCellMar cellMar = pr.isSetTblCellMar() ? pr.getTblCellMar() : pr.addNewTblCellMar();
        setCellMargin(cellMar.isSetTop() ? cellMar.getTop() : cellMar.addNewTop());
        setCellMargin(cellMar.isSetLeft() ? cellMar.getLeft() : cellMar.addNewLeft());
        setCellMargin(cellMar.isSetBottom() ? cellMar.getBottom() : cellMar.addNewBottom());
        setCellMargin(cellMar.isSetRight() ? cellMar.getRight() : cellMar.addNewRight());
    }

    private static void setCellMargin(CTTblWidth width) {
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(20));
    }

    private static void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
        border.setColor("auto");
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        setHeaderCellText(cell, text, ParagraphAlignment.CENTER);
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text, ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);
        run.setBold(false);
        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setHeaderCellPageCount(XWPFTableCell cell, ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);
        run.setText("Количество страниц: ");
        appendField(paragraph, "PAGE");
        XWPFRun separator = paragraph.createRun();
        separator.setFontFamily(FONT_NAME);
        separator.setFontSize(TABLE_FONT_SIZE);
        separator.setText(" / ");
        appendField(paragraph, "NUMPAGES");
    }

    private static void appendField(XWPFParagraph paragraph, String instr) {
        XWPFRun runBegin = paragraph.createRun();
        runBegin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun runInstr = paragraph.createRun();
        runInstr.getCTR().addNewInstrText().setStringValue(instr);

        XWPFRun runSep = paragraph.createRun();
        runSep.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun runText = paragraph.createRun();
        runText.setText("1");

        XWPFRun runEnd = paragraph.createRun();
        runEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void stretchTableToFullWidth(XWPFTable table) {
        table.setWidthType(TableWidthType.PCT);
        table.setWidth("100%");
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
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
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

    private static void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            if (cell == null) {
                continue;
            }
            if (cell.getCTTc().getTcPr() == null) {
                cell.getCTTc().addNewTcPr();
            }
            if (colIndex == fromCol) {
                cell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().getTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    private static void addParagraphText(XWPFDocument document,
                                         String text,
                                         int fontSize,
                                         ParagraphAlignment alignment,
                                         boolean bold) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(alignment);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
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

    private static void addSpacer(XWPFDocument document, int lines) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        for (int i = 0; i < lines; i++) {
            run.addBreak();
        }
    }

    private static void tightenHeaderRowHeights(XWPFTable table) {
        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if (row != null) {
                row.setHeight(360);
            }
        }
    }
}
