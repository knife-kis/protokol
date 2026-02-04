package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.export.PrintSetupUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class RadiationJournalExporter {
    private static final String JOURNAL_FILE_NAME = "Журнал.xlsx";
    private static final int DEFAULT_POINT_COUNT = 1;
    private static final int START_SK13_NUMBER = 17;
    private static final int TITLE_TOP_OFFSET = 2;
    private static final double JOURNAL_COLUMN_SCALE = 0.25;

    private RadiationJournalExporter() {
    }

    static void generate(File sourceFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveJournalFile(mapFile);
        ProtocolData data = resolveProtocolData(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Журнал");
            applySheetDefaults(sheet);
            applyHeader(sheet);

            Styles styles = new Styles(workbook);
            buildJournalSheet(sheet, styles, data);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                workbook.write(out);
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

    private static void applySheetDefaults(Sheet sheet) {
        PrintSetup setup = sheet.getPrintSetup();
        setup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        setup.setLandscape(true);
        setup.setFitWidth((short) 1);
        setup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        PrintSetupUtils.applyDuplexShortEdge(sheet);

        double[] widths = new double[]{14, 18, 14, 14, 14, 14, 14, 18, 18};
        for (int col = 0; col < widths.length; col++) {
            double scaled = widths[col] * JOURNAL_COLUMN_SCALE;
            sheet.setColumnWidth(col, (int) Math.round(scaled * 256));
        }
        sheet.setDefaultRowHeightInPoints(15);
    }

    private static void applyHeader(Sheet sheet) {
        Header header = sheet.getHeader();
        String font = "&\"Arial\"&9";
        header.setLeft(font + "Испытательная лаборатория\nООО «ЦИТ»");
        header.setCenter(font + "Журнал регистрации результатов\nопределения плотности потока радона\nФ6 РИ ИЛ 2-2023");
        header.setRight(font + "\nКоличество страниц: &[Страница] / &[Страниц] \n ");
    }

    private static void buildJournalSheet(Sheet sheet, Styles styles, ProtocolData data) {
        int row = TITLE_TOP_OFFSET;

        mergeWithValue(sheet, row, row + 3, 0, 8,
                "ЖУРНАЛ\nРегистрации результатов определения плотности потока радона\n№5/2023",
                styles.title28, true);
        for (int r = row; r <= row + 3; r++) {
            ensureRow(sheet, r).setHeightInPoints(32);
        }
        row += 4;

        row += 3;
        mergeWithValue(sheet, row, row + 1, 0, 8,
                "Начат: _______________________\nОкончен: ____________________\nСрок хранения в архиве: 3 года.",
                styles.text10Left, false);
        ensureRow(sheet, row).setHeightInPoints(36);
        ensureRow(sheet, row + 1).setHeightInPoints(36);
        row += 2;

        sheet.setRowBreak(row);
        row++;

        int tableStart = row;
        for (int i = 0; i < 13; i++) {
            ensureRow(sheet, tableStart + i);
        }

        mergeWithValue(sheet, tableStart, tableStart, 0, 5,
                "Ответственный за ведение журнала, Должность, Фамилия, Инициалы.",
                styles.table9Header, true);
        mergeWithValue(sheet, tableStart, tableStart, 6, 8,
                "Образец подписи", styles.table9Header, true);

        String preparedLine = data.preparedLine;
        mergeWithValue(sheet, tableStart + 1, tableStart + 1, 0, 5,
                preparedLine, styles.table9Left, true);
        mergeWithValue(sheet, tableStart + 1, tableStart + 1, 6, 8,
                "", styles.table9Center, true);

        mergeWithValue(sheet, tableStart + 2, tableStart + 2, 0, 5,
                "Допущены к ведению журнала, Должность, Фамилия, Инициалы.",
                styles.table9Header, true);
        mergeWithValue(sheet, tableStart + 2, tableStart + 2, 6, 8,
                "Образец подписи", styles.table9Header, true);

        mergeWithValue(sheet, tableStart + 3, tableStart + 3, 0, 5,
                preparedLine, styles.table9Left, true);
        mergeWithValue(sheet, tableStart + 3, tableStart + 3, 6, 8,
                "", styles.table9Center, true);

        String tarnoLine = "Заведующий лабораторией Тарновский М.О.";
        String row5Text = preparedLine.toLowerCase(Locale.ROOT).contains("тарновский") ? "" : tarnoLine;
        mergeWithValue(sheet, tableStart + 4, tableStart + 4, 0, 5,
                row5Text, styles.table9Left, true);
        mergeWithValue(sheet, tableStart + 4, tableStart + 4, 6, 8,
                "", styles.table9Center, true);

        for (int i = 5; i <= 5; i++) {
            mergeWithValue(sheet, tableStart + i, tableStart + i, 0, 5,
                    "", styles.table9Left, true);
            mergeWithValue(sheet, tableStart + i, tableStart + i, 6, 8,
                    "", styles.table9Center, true);
        }

        mergeWithValue(sheet, tableStart + 6, tableStart + 6, 0, 8,
                "Список сокращений (при необходимости)", styles.table9Header, true);

        mergeWithValue(sheet, tableStart + 7, tableStart + 7, 0, 5,
                "Условное обозначение", styles.table9Header, true);
        mergeWithValue(sheet, tableStart + 7, tableStart + 7, 6, 8,
                "Расшифровка условного обозначения", styles.table9Header, true);

        for (int i = 8; i < 13; i++) {
            mergeWithValue(sheet, tableStart + i, tableStart + i, 0, 5,
                    "", styles.table9Left, true);
            mergeWithValue(sheet, tableStart + i, tableStart + i, 6, 8,
                    "", styles.table9Center, true);
        }

        row = tableStart + 13;
        mergeWithValue(sheet, row, row, 0, 8,
                "Регистрация результатов проверок ведения журнала", styles.text10Center, false);
        row++;

        int checkTableStart = row;
        for (int i = 0; i < 4; i++) {
            ensureRow(sheet, checkTableStart + i);
        }
        String[] checkHeaders = new String[]{
                "Дата проверки",
                "Период проверенных записей",
                "Результат проверки",
                "Ф.И.О. проверяющего, подпись",
                "Ознакомление ответственного",
                "Отметка об устранении"
        };
        for (int col = 0; col < 6; col++) {
            setCell(sheet, checkTableStart, col, checkHeaders[col], styles.table9Header);
        }
        for (int r = 1; r < 4; r++) {
            for (int c = 0; c < 6; c++) {
                setCell(sheet, checkTableStart + r, c, "", styles.table9Center);
            }
        }
        addRegionBorders(sheet, checkTableStart, checkTableStart + 3, 0, 5);

        row = checkTableStart + 4;
        sheet.setRowBreak(row);
        row++;

        int infoTableRow = row;
        ensureRow(sheet, infoTableRow);
        mergeWithValue(sheet, infoTableRow, infoTableRow, 0, 1,
                "Период наблюдений:", styles.table9Header, true);
        mergeWithValue(sheet, infoTableRow, infoTableRow, 2, 4,
                data.observationPeriod, styles.table9Left, true);
        mergeWithValue(sheet, infoTableRow, infoTableRow, 5, 6,
                "Населенный пункт:", styles.table9Header, true);
        mergeWithValue(sheet, infoTableRow, infoTableRow, 7, 8,
                data.settlement, styles.table9Left, true);

        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Методика: " + data.method, styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Наименование заказчика: " + data.customerName, styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Заявка заказчика: " + data.customerRequest, styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Присутствие представителя Заказчика при отборе: " + data.customerPresence,
                styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Подготовка к отбору пробы:", styles.text10Left, false);
        row++;

        int prepTableStart = row;
        for (int i = 0; i < 4; i++) {
            ensureRow(sheet, prepTableStart + i);
        }
        String[] prepHeaders = new String[]{
                "Дата регенерации",
                "Продолжительность регенерации",
                "t активированного угля, °С",
                "Помещение угля в СК-13 (+/-)",
                "Уголь пригоден для отбора (+)/не пригоден для отбора (-)",
                "Подпись ответственного ха подготовку к отбору пробы"
        };
        for (int col = 0; col < prepHeaders.length; col++) {
            setCell(sheet, prepTableStart, col, prepHeaders[col], styles.table9Header);
        }
        for (int r = 1; r < 4; r++) {
            for (int c = 0; c < prepHeaders.length; c++) {
                setCell(sheet, prepTableStart + r, c, "", styles.table9Center);
            }
        }
        addRegionBorders(sheet, prepTableStart, prepTableStart + 3, 0, 5);
        row = prepTableStart + 4;

        mergeWithValue(sheet, row, row, 0, 8,
                "Условные обозначения: + выполнено;  - не выполнено. ",
                styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Требования к условиям окружающей среды:", styles.text10Left, false);
        row++;

        int envTableStart = row;
        for (int i = 0; i < 10; i++) {
            ensureRow(sheet, envTableStart + i);
        }

        mergeWithValue(sheet, envTableStart, envTableStart + 1, 0, 0,
                "№ п/п", styles.table9Header, true);
        mergeWithValue(sheet, envTableStart, envTableStart + 1, 1, 1,
                "Шифр методики/ описание деятельности", styles.table9Header, true);
        mergeWithValue(sheet, envTableStart, envTableStart, 2, 6,
                "Рабочие условия", styles.table9Header, true);
        mergeWithValue(sheet, envTableStart, envTableStart + 1, 7, 8,
                "Дополнительные сведения", styles.table9Header, true);

        setCell(sheet, envTableStart + 1, 2, "Т, °С", styles.table9Header);
        setCell(sheet, envTableStart + 1, 3, "Р, мм рт. ст", styles.table9Header);
        setCell(sheet, envTableStart + 1, 4, "Относительная влажность, %", styles.table9Header);
        setCell(sheet, envTableStart + 1, 5, "Напряжение сети, В", styles.table9Header);
        setCell(sheet, envTableStart + 1, 6, "Частота тока, Гц", styles.table9Header);

        mergeWithValue(sheet, envTableStart + 2, envTableStart + 5, 0, 0,
                "1.", styles.table9Center, true);
        mergeWithValue(sheet, envTableStart + 2, envTableStart + 5, 1, 1,
                "Методика измерений плотности потока радона с поверхности земли и " +
                        "строительных конструкций (свидетельство № 40090.6К816 об аттестации МВИ) " +
                        "при отборе образцов",
                styles.table9Left, true);
        mergeWithValue(sheet, envTableStart + 2, envTableStart + 2, 2, 6,
                "НОРМА", styles.table9Header, true);
        setCell(sheet, envTableStart + 3, 2, "от минус 15°С до +40 °С", styles.table9Center);
        setCell(sheet, envTableStart + 3, 3, "-", styles.table9Center);
        setCell(sheet, envTableStart + 3, 4, "-", styles.table9Center);
        setCell(sheet, envTableStart + 3, 5, "-", styles.table9Center);
        setCell(sheet, envTableStart + 3, 6, "-", styles.table9Center);
        mergeWithValue(sheet, envTableStart + 4, envTableStart + 4, 2, 6,
                "Фактически", styles.table9Header, true);
        for (int c = 2; c <= 6; c++) {
            setCell(sheet, envTableStart + 5, c, "", styles.table9Center);
        }
        mergeWithValue(sheet, envTableStart + 2, envTableStart + 5, 7, 8,
                "Контроль осуществляется в течение времени пассивного отбора образцов, " +
                        "для чего используется сигнализатор, который световым сигналом оповещает " +
                        "работника ИЛ о выходе за пределы температурного \nрежима, которые влияют " +
                        "на погрешность измерения. ",
                styles.table9Left, true);

        mergeWithValue(sheet, envTableStart + 6, envTableStart + 9, 0, 0,
                "2.", styles.table9Center, true);
        mergeWithValue(sheet, envTableStart + 6, envTableStart + 9, 1, 1,
                "Проведение измерений по Методике измерений плотности потока радона с " +
                        "поверхности земли и строительных конструкций (свидетельство № 40090.6К816 об аттестации МВИ)",
                styles.table9Left, true);
        mergeWithValue(sheet, envTableStart + 6, envTableStart + 6, 2, 6,
                "НОРМА", styles.table9Header, true);
        setCell(sheet, envTableStart + 7, 2, "от плюс 10°С до плюс 35 °С", styles.table9Center);
        setCell(sheet, envTableStart + 7, 3, "от 630 мм рт. ст. до 802 мм рт. ст.", styles.table9Center);
        setCell(sheet, envTableStart + 7, 4, "до 95 % при температуре плюс 30 °С", styles.table9Center);
        setCell(sheet, envTableStart + 7, 5, "220+22 -33", styles.table9Center);
        setCell(sheet, envTableStart + 7, 6, "50±1", styles.table9Center);
        mergeWithValue(sheet, envTableStart + 8, envTableStart + 8, 2, 6,
                "Фактически", styles.table9Header, true);
        for (int c = 2; c <= 6; c++) {
            setCell(sheet, envTableStart + 9, c, "", styles.table9Center);
        }
        mergeWithValue(sheet, envTableStart + 6, envTableStart + 9, 7, 8,
                "Руководство по эксплуатации  \nФМКТ.136132.134 РЭ \n" +
                        "Комплекс измерительный для мониторинга радона «КАМЕРА-01» ",
                styles.table9Left, true);

        addRegionBorders(sheet, envTableStart, envTableStart + 9, 0, 8);

        row = envTableStart + 10;
        sheet.setRowBreak(row);
        row++;

        mergeWithValue(sheet, row, row, 0, 8,
                "Сведения об отборе:", styles.text10Left, false);
        row++;

        int sampleTableStart = row;
        int pointCount = Math.max(DEFAULT_POINT_COUNT, data.pointCount);
        int sampleTableRows = 2 + pointCount;
        for (int i = 0; i < sampleTableRows; i++) {
            ensureRow(sheet, sampleTableStart + i);
        }

        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 0, 0,
                "Местоположение пункта наблюдения", styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 1, 1,
                "Подготовка участка (очистка от дерна, рыхление земли, перед установкой НК-32 прошло более 1 ч): +/-",
                styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 2, 2,
                "Номер СК-13", styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 3, 3,
                "Количество совместно экспонированных НК-32", styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 4, 4,
                "Место установки НК-32", styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart, 5, 6,
                "Дата и время экспонирования (с точностью до 1 мин)", styles.table9Header, true);
        setCell(sheet, sampleTableStart + 1, 5, "Начала", styles.table9Header);
        setCell(sheet, sampleTableStart + 1, 6, "Окончания", styles.table9Header);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 7, 7,
                "Отклонения, дополнения или исключения от условий отбора",
                styles.table9Header, true);
        mergeWithValue(sheet, sampleTableStart, sampleTableStart + 1, 8, 8,
                "Пригоден образец после транспортировки  для последующих измерений: ДА/НЕТ",
                styles.table9Header, true);

        if (pointCount > 0) {
            mergeWithValue(sheet, sampleTableStart + 2, sampleTableStart + 1 + pointCount, 0, 0,
                    data.observationLocation, styles.table9Left, true);
        }

        int seq = START_SK13_NUMBER;
        for (int i = 0; i < pointCount; i++) {
            int dataRow = sampleTableStart + 2 + i;
            setCell(sheet, dataRow, 1, "", styles.table9Center);
            setCell(sheet, dataRow, 2, String.format("%04d", seq++), styles.table9Center);
            setCell(sheet, dataRow, 3, "1", styles.table9Center);
            setCell(sheet, dataRow, 4, "т" + (i + 1), styles.table9Center);
            setCell(sheet, dataRow, 5, "", styles.table9Center);
            setCell(sheet, dataRow, 6, "", styles.table9Center);
            setCell(sheet, dataRow, 7, "", styles.table9Center);
            setCell(sheet, dataRow, 8, "", styles.table9Center);
        }
        addRegionBorders(sheet, sampleTableStart, sampleTableStart + sampleTableRows - 1, 0, 8);

        row = sampleTableStart + sampleTableRows;
        mergeWithValue(sheet, row, row, 0, 8,
                "Условные обозначения: + выполнено; - не выполнено.", styles.text10Left, false);
        row++;
        mergeWithValue(sheet, row, row, 0, 8,
                "Отбор произвел: ", styles.text10Left, false);
        row++;

        int takerTableStart = row;
        for (int i = 0; i < 2; i++) {
            ensureRow(sheet, takerTableStart + i);
        }
        for (int c = 0; c < 4; c++) {
            setCell(sheet, takerTableStart, c, "", styles.table9Center);
        }
        String[] takerHeaders = new String[]{"Должность", "подпись", "Ф.И.О", "дата"};
        for (int c = 0; c < 4; c++) {
            setCell(sheet, takerTableStart + 1, c, takerHeaders[c], styles.table9Header);
        }
        addRegionBorders(sheet, takerTableStart, takerTableStart + 1, 0, 3);
        row = takerTableStart + 2;

        mergeWithValue(sheet, row, row, 0, 8,
                "На измерения образец принял: ", styles.text10Left, false);
        row++;

        int acceptTableStart = row;
        for (int i = 0; i < 2; i++) {
            ensureRow(sheet, acceptTableStart + i);
        }
        for (int c = 0; c < 4; c++) {
            setCell(sheet, acceptTableStart, c, "", styles.table9Center);
        }
        for (int c = 0; c < 4; c++) {
            setCell(sheet, acceptTableStart + 1, c, takerHeaders[c], styles.table9Header);
        }
        addRegionBorders(sheet, acceptTableStart, acceptTableStart + 1, 0, 3);
        row = acceptTableStart + 2;

        mergeWithValue(sheet, row, row, 0, 8,
                "Измерения:", styles.text10Left, false);
        row++;

        int measurementTableStart = row;
        String[] measurementHeaders = new String[]{
                "Номер СК-13",
                "Дата и время измерения",
                "Значение плотности потока радона, мБк/м^2с",
                "Погрешность измерения, мБк/м^2с",
                "Подпись измерителя"
        };
        ensureRow(sheet, measurementTableStart);
        for (int c = 0; c < measurementHeaders.length; c++) {
            setCell(sheet, measurementTableStart, c, measurementHeaders[c], styles.table9Header);
        }
        for (int i = 0; i < pointCount; i++) {
            int dataRow = measurementTableStart + 1 + i;
            ensureRow(sheet, dataRow);
            for (int c = 0; c < measurementHeaders.length; c++) {
                setCell(sheet, dataRow, c, "", styles.table9Center);
            }
        }
        addRegionBorders(sheet, measurementTableStart, measurementTableStart + pointCount, 0, 4);

        row = measurementTableStart + pointCount + 1;
        sheet.setRowBreak(row);
        row++;

        mergeWithValue(sheet, row, row, 0, 8,
                "ЛИСТ РАССЫЛКИ ДОКУМЕНТОВ", styles.text10Center, false);
        row++;

        int distributionTableStart = row;
        for (int i = 0; i < 4; i++) {
            ensureRow(sheet, distributionTableStart + i);
        }
        String[] distributionHeaders = new String[]{
                "Номер учтенной копии",
                "Ф.И.О., должность",
                "Подпись о получении учтенной копии, дата",
                "Отметка об изъятии у получателя учтенной копии, уничтожении учтенной копии: " +
                        "подпись уполномоченного работника ИЛ, дата"
        };
        for (int c = 0; c < distributionHeaders.length; c++) {
            setCell(sheet, distributionTableStart, c, distributionHeaders[c], styles.table9Header);
        }
        for (int r = 1; r < 4; r++) {
            for (int c = 0; c < distributionHeaders.length; c++) {
                setCell(sheet, distributionTableStart + r, c, "", styles.table9Center);
            }
        }
        addRegionBorders(sheet, distributionTableStart, distributionTableStart + 3, 0, 3);

        row = distributionTableStart + 4;
        sheet.setRowBreak(row);
        row++;

        mergeWithValue(sheet, row, row, 0, 8,
                "ЛИСТ РЕГИСТРАЦИИ ИЗМЕНЕНИЙ", styles.text10Center, false);
        row++;

        int changeTableStart = row;
        for (int i = 0; i < 6; i++) {
            ensureRow(sheet, changeTableStart + i);
        }

        mergeWithValue(sheet, changeTableStart, changeTableStart + 1, 0, 0,
                "№\nизменения ", styles.table9Header, true);
        mergeWithValue(sheet, changeTableStart, changeTableStart, 1, 4,
                "Номер листа (страницы), пункта", styles.table9Header, true);
        mergeWithValue(sheet, changeTableStart, changeTableStart + 1, 5, 5,
                "Дата \nутверждения\n", styles.table9Header, true);
        mergeWithValue(sheet, changeTableStart, changeTableStart + 1, 6, 6,
                "Ф.И.О., должность лица, внесшего изменения", styles.table9Header, true);
        mergeWithValue(sheet, changeTableStart, changeTableStart + 1, 7, 7,
                "Подпись МК", styles.table9Header, true);

        setCell(sheet, changeTableStart + 1, 1, "измененного", styles.table9Header);
        setCell(sheet, changeTableStart + 1, 2, "замененного", styles.table9Header);
        setCell(sheet, changeTableStart + 1, 3, "нового", styles.table9Header);
        setCell(sheet, changeTableStart + 1, 4, "аннулированного", styles.table9Header);

        for (int r = 2; r < 6; r++) {
            for (int c = 0; c < 8; c++) {
                if (getCell(sheet, changeTableStart + r, c) == null) {
                    setCell(sheet, changeTableStart + r, c, "", styles.table9Center);
                }
            }
        }
        addRegionBorders(sheet, changeTableStart, changeTableStart + 5, 0, 7);
    }

    private static Cell getCell(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }
        return row.getCell(colIndex);
    }

    private static Row ensureRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private static void setCell(Sheet sheet, int rowIndex, int colIndex, String value, CellStyle style) {
        Row row = ensureRow(sheet, rowIndex);
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value == null ? "" : value);
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private static void mergeWithValue(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol,
                                       String value, CellStyle style, boolean borders) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);
        for (int r = firstRow; r <= lastRow; r++) {
            for (int c = firstCol; c <= lastCol; c++) {
                setCell(sheet, r, c, "", style);
            }
        }
        setCell(sheet, firstRow, firstCol, value, style);
        if (borders) {
            applyBorders(sheet, region);
        }
    }

    private static void applyBorders(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private static void addRegionBorders(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        applyBorders(sheet, region);
    }

    private static ProtocolData resolveProtocolData(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return ProtocolData.empty();
        }
        String preparedLine = "";
        String observationPeriod = "";
        String settlement = "";
        String method = "";
        String customerName = "";
        String customerRequest = "";
        String customerPresence = "";
        String observationLocation = "";
        int pointCount = DEFAULT_POINT_COUNT;

        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {

            preparedLine = resolvePreparedLine(workbook);
            observationPeriod = resolveObservationPeriod(workbook);
            settlement = resolveSettlement(workbook);
            method = resolveMethod(workbook);
            customerName = resolveValueByPrefix(workbook,
                    "Наименование и контактные данные заявителя (заказчика):");
            customerRequest = resolveCustomerRequest(workbook);
            customerPresence = resolveValueByPrefix(workbook,
                    "Измерения проводились в присутствии представителя заказчика:");
            observationLocation = resolveValueByPrefix(workbook,
                    "Наименование предприятия, организации, объекта, где производились измерения:");
            pointCount = Math.max(DEFAULT_POINT_COUNT, countSamplingPoints(workbook));
        } catch (Exception ignored) {
            return ProtocolData.empty();
        }

        return new ProtocolData(
                preparedLine,
                observationPeriod,
                settlement,
                method,
                customerName,
                customerRequest,
                customerPresence,
                observationLocation,
                pointCount
        );
    }

    private static String resolvePreparedLine(Workbook workbook) {
        String defaultValue = "Заведующий лабораторией Тарновский М.О.";
        Sheet sheet = findSheet(workbook, "ППР");
        if (sheet == null) {
            return defaultValue;
        }
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            StringBuilder rowText = new StringBuilder();
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
            }
            String mergedText = rowText.toString();
            if (!mergedText.toLowerCase(Locale.ROOT).contains("протокол подготовил")) {
                continue;
            }
            String lower = mergedText.toLowerCase(Locale.ROOT);
            if (lower.contains("тарновский")) {
                return "Заведующий лабораторией Тарновский М.О.";
            }
            if (lower.contains("белов")) {
                return "Инженер Белов Д.А.";
            }
            return defaultValue;
        }
        return defaultValue;
    }

    private static String resolveObservationPeriod(Workbook workbook) {
        String text = resolveValueByPrefix(workbook,
                "Дополнительные сведения (характеристика объекта): Измерения были проведены");
        if (text.isBlank()) {
            return "";
        }
        List<String> dates = new ArrayList<>();
        Matcher matcher = Pattern.compile("\\d{2}\\.\\d{2}\\.\\d{4}").matcher(text);
        while (matcher.find()) {
            dates.add(matcher.group());
        }
        if (!dates.isEmpty()) {
            List<String> uniqueDates = new ArrayList<>();
            for (String date : dates) {
                if (!uniqueDates.contains(date)) {
                    uniqueDates.add(date);
                }
            }
            return String.join("; ", uniqueDates);
        }
        return text.replace(",", ";");
    }

    private static String resolveSettlement(Workbook workbook) {
        String address = resolveValueByPrefix(workbook, "Адрес предприятия (объекта):");
        if (address.isBlank()) {
            return "";
        }
        int comma = address.indexOf(',');
        return comma >= 0 ? address.substring(0, comma).trim() : address.trim();
    }

    private static String resolveMethod(Workbook workbook) {
        Sheet sheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : null;
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            for (Cell cell : row) {
                String value = formatter.formatCellValue(cell, evaluator).trim();
                if (value.contains("Мощность дозы гамма-излучения")) {
                    int nextCol = cell.getColumnIndex() + 1;
                    String nextValue = readMergedCellValue(sheet, row.getRowNum(), nextCol, formatter, evaluator);
                    if (!nextValue.isBlank()) {
                        return nextValue;
                    }
                }
            }
        }
        return "";
    }

    private static String resolveCustomerRequest(Workbook workbook) {
        String basis = resolveValueByPrefix(workbook, "Основание для измерений:");
        if (basis.isBlank()) {
            return "";
        }
        int idx = basis.toLowerCase(Locale.ROOT).indexOf("заявка");
        if (idx >= 0 && idx + "заявка".length() < basis.length()) {
            return basis.substring(idx + "заявка".length()).trim();
        }
        return basis;
    }

    private static int countSamplingPoints(Workbook workbook) {
        Sheet sheet = findSheet(workbook, "ППР");
        if (sheet == null) {
            return 0;
        }
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int startRow = 4;
        String startText = readMergedCellValue(sheet, startRow, 1, formatter, evaluator).toLowerCase(Locale.ROOT);
        if (!startText.contains("точка")) {
            int last = sheet.getLastRowNum();
            for (int r = 0; r <= Math.min(last, 200); r++) {
                String text = readMergedCellValue(sheet, r, 1, formatter, evaluator).toLowerCase(Locale.ROOT);
                if (text.contains("точка")) {
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

    private static String resolveValueByPrefix(Workbook workbook, String prefix) {
        if (workbook == null || prefix == null) {
            return "";
        }
        if (workbook.getNumberOfSheets() == 0) {
            return "";
        }
        Sheet sheet = workbook.getSheetAt(0);
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell, evaluator).trim();
                if (text.startsWith(prefix)) {
                    String tail = text.substring(prefix.length()).trim();
                    if (!tail.isBlank()) {
                        return trimLeadingPunctuation(tail);
                    }
                    int nextCol = cell.getColumnIndex() + 1;
                    String nextValue = readMergedCellValue(sheet, row.getRowNum(), nextCol, formatter, evaluator);
                    if (!nextValue.isBlank()) {
                        return trimLeadingPunctuation(nextValue);
                    }
                }
            }
        }
        return "";
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
            Cell cell = row.getCell(region.getFirstColumn());
            return cell == null ? "" : formatter.formatCellValue(cell, evaluator).trim();
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return cell == null ? "" : formatter.formatCellValue(cell, evaluator).trim();
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String trimLeadingPunctuation(String text) {
        if (text == null) {
            return "";
        }
        int index = 0;
        while (index < text.length()) {
            char ch = text.charAt(index);
            if (Character.isLetterOrDigit(ch)) {
                break;
            }
            if (!Character.isWhitespace(ch)) {
                index++;
                continue;
            }
            index++;
        }
        return text.substring(index).trim();
    }

    private record ProtocolData(String preparedLine,
                                String observationPeriod,
                                String settlement,
                                String method,
                                String customerName,
                                String customerRequest,
                                String customerPresence,
                                String observationLocation,
                                int pointCount) {
        static ProtocolData empty() {
            return new ProtocolData(
                    "Заведующий лабораторией Тарновский М.О.",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    "",
                    DEFAULT_POINT_COUNT
            );
        }
    }

    private static class Styles {
        final CellStyle text10Left;
        final CellStyle text10Center;
        final CellStyle table9Header;
        final CellStyle table9Center;
        final CellStyle table9Left;
        final CellStyle title28;

        Styles(Workbook workbook) {
            Font font10 = workbook.createFont();
            font10.setFontName("Arial");
            font10.setFontHeightInPoints((short) 10);

            Font font9 = workbook.createFont();
            font9.setFontName("Arial");
            font9.setFontHeightInPoints((short) 9);

            Font font9Bold = workbook.createFont();
            font9Bold.setFontName("Arial");
            font9Bold.setFontHeightInPoints((short) 9);
            font9Bold.setBold(true);

            Font font28 = workbook.createFont();
            font28.setFontName("Arial");
            font28.setFontHeightInPoints((short) 28);
            font28.setBold(true);

            text10Left = createStyle(workbook, font10, HorizontalAlignment.LEFT, false, false);
            text10Center = createStyle(workbook, font10, HorizontalAlignment.CENTER, false, false);
            table9Header = createStyle(workbook, font9Bold, HorizontalAlignment.CENTER, true, true);
            table9Center = createStyle(workbook, font9, HorizontalAlignment.CENTER, true, true);
            table9Left = createStyle(workbook, font9, HorizontalAlignment.LEFT, true, true);
            title28 = createStyle(workbook, font28, HorizontalAlignment.CENTER, true, true);
        }

        private static CellStyle createStyle(Workbook workbook,
                                             Font font,
                                             HorizontalAlignment alignment,
                                             boolean wrap,
                                             boolean withBorders) {
            CellStyle style = workbook.createCellStyle();
            style.setFont(font);
            style.setAlignment(alignment);
            style.setVerticalAlignment(VerticalAlignment.CENTER);
            style.setWrapText(wrap);
            if (withBorders) {
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
            }
            return style;
        }
    }
}
