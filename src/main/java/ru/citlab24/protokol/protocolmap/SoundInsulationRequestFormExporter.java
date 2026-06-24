package ru.citlab24.protokol.protocolmap;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Pattern;
import java.util.regex.Matcher;

final class SoundInsulationRequestFormExporter {
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int PLAN_TABLE_FONT_SIZE = 10;
    private static final double REQUEST_TABLE_WIDTH_SCALE = 0.8;
    private static final String PLAN_TABLE_HEADER = "Наименование объекта измерений";
    private static final String CUSTOMER_SECTION_HEADER =
            "Данные, предоставленные Заказчиком, за которые он несет ответственность";
    private static final String PLAN_APPENDIX_TITLE = "ПЛАН ИЗМЕРЕНИЙ";
    private static final String PLAN_APPENDIX_HEADER = "Приложение – План измерений";
    private static final String CUSTOMER_APPENDIX_HEADER = "Приложение – Данные, предоставлены Заказчиком";
    private static final String CUSTOMER_DATA_START =
            "Объект испытаний – внутренние ограждающие конструкции помещений, " +
                    "их монтаж осуществлен заказчиком согласно требованиям технической документации.";
    private static final String CUSTOMER_DATA_END =
            "14. Сведения о дополнении, отклонении или исключении из методов: -";
    private static final String ROOM_PARAMS_START = "Параметры помещений и испытываемой поверхности:";
    private static final String ROOM_PARAMS_END = "17. Результаты измерений";
    private static final String ROOM_PARAMS_ALT_START = "16. Параметры помещений и испытываемой поверхности:";
    private static final String CUSTOMER_INFO_LABEL =
            "1. Наименование и контактные данные заявителя (заказчика):";
    private static final String AREA_BETWEEN_ROOMS_MARKER =
            "Площадь испытываемой поверхности между помещениями";
    private static final String NORMATIVE_REQUIREMENTS_WITH_BRACKET_MARKER = "(нормативные требования";
    private static final String LNW_SENTENCE_MARKER =
            "Индекс приведенного уровня ударного шума (Lnw) по результатам измерений для перекрытия между помещениями";
    private static final String RW_FLOOR_SENTENCE_MARKER =
            "Индекс изоляции воздушного шума (Rw) по результатам измерений для перекрытия между помещениями";
    private static final String RW_PARTITION_SENTENCE_MARKER =
            "Индекс изоляции воздушного шума (Rw) по результатам измерений для перегородки между помещениями";
    private static final String APPLICATION_BASIS_LABEL = "6. Основание для измерений";
    private static final String APPLICATION_BASIS_ALT_LABEL = "Основание для измерений";
    private static final Pattern EMAIL_PATTERN =
            Pattern.compile("[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}");
    private static final Pattern PHONE_PREFIX_PATTERN =
            Pattern.compile("(?iu)^(?:контактный\\s+телефон|телефон|тел|т)\\.?\\s*[:.]?\\s*");

    private SoundInsulationRequestFormExporter() {
    }

    static void generate(File protocolFile, File mapFile, String workDeadline, String customerInn) {
        generate(protocolFile, mapFile, workDeadline, customerInn, true);
    }

    static void generateAppendix(File protocolFile, File mapFile, String workDeadline, String customerInn) {
        generate(protocolFile, mapFile, workDeadline, customerInn, false);
    }

    private static void generate(File protocolFile,
                                 File mapFile,
                                 String workDeadline,
                                 String customerInn,
                                 boolean includeRequestSection) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveRequestFormFile(mapFile);
        String applicationNumber = resolveApplicationNumber(protocolFile, mapFile);
        CustomerInfo customerInfo = extractCustomerInfo(protocolFile);
        File planFile = SoundInsulationMeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
        List<List<String>> planRows = extractPlanRows(planFile);
        List<String> customerLines = extractCustomerDataLines(protocolFile);
        List<List<String>> roomParamsRows = extractRoomParamsTable(protocolFile);
        List<String> roomParamsLines = roomParamsRows.isEmpty() ? extractRoomParamsLines(protocolFile) : List.of();
        String areaBetweenRoomsLine = extractAreaBetweenRoomsLine(protocolFile);
        List<String> measurementRequirementLines = extractMeasurementRequirementLines(protocolFile);

        if (planRows.isEmpty()) {
            List<String> empty = new ArrayList<>();
            for (int i = 0; i < 5; i++) {
                empty.add("");
            }
            planRows.add(empty);
        }
        applyPlanDeadline(planRows, workDeadline);

        try (XWPFDocument document = new XWPFDocument()) {
            RequestFormExporter.applyStandardHeader(document);

            if (includeRequestSection) {
                addRequestSection(document, applicationNumber, customerInfo, customerInn, workDeadline);
            }

            XWPFParagraph appendixHeader = document.createParagraph();
            appendixHeader.setAlignment(ParagraphAlignment.RIGHT);
            setParagraphSpacing(appendixHeader);
            XWPFRun appendixHeaderRun = appendixHeader.createRun();
            appendixHeaderRun.setFontFamily(FONT_NAME);
            appendixHeaderRun.setFontSize(FONT_SIZE);
            setRunTextWithBreaks(appendixHeaderRun,
                    "Приложения к заявке № " + applicationNumber + "\n" + PLAN_APPENDIX_HEADER);

            XWPFParagraph appendixTitle = document.createParagraph();
            appendixTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(appendixTitle);
            XWPFRun appendixTitleRun = appendixTitle.createRun();
            appendixTitleRun.setText(PLAN_APPENDIX_TITLE);
            appendixTitleRun.setFontFamily(FONT_NAME);
            appendixTitleRun.setFontSize(FONT_SIZE);
            appendixTitleRun.setBold(true);

            XWPFTable planTable = document.createTable(planRows.size(), 5);
            configureTableLayout(planTable, new int[]{2500, 2500, 2100, 2560, 2900});
            for (int rowIndex = 0; rowIndex < planRows.size(); rowIndex++) {
                List<String> row = planRows.get(rowIndex);
                boolean isHeader = rowIndex == 0;
                ParagraphAlignment alignment = isHeader ? ParagraphAlignment.CENTER : ParagraphAlignment.LEFT;
                for (int colIndex = 0; colIndex < 5; colIndex++) {
                    String value = colIndex < row.size() ? row.get(colIndex) : "";
                    setTableCellText(planTable.getRow(rowIndex).getCell(colIndex), value,
                            PLAN_TABLE_FONT_SIZE, isHeader, alignment);
                }
            }

            XWPFParagraph spacerAfterPlan = document.createParagraph();
            setParagraphSpacing(spacerAfterPlan);

            addParagraphWithLineBreaks(document,
                    "Представитель заказчика _______________________________________________\n" +
                            "                                                 (Должность, ФИО, контактные данные)  ");

            addPageBreak(document);

            XWPFParagraph customerAppendixHeader = document.createParagraph();
            customerAppendixHeader.setAlignment(ParagraphAlignment.RIGHT);
            setParagraphSpacing(customerAppendixHeader);
            XWPFRun customerAppendixHeaderRun = customerAppendixHeader.createRun();
            customerAppendixHeaderRun.setFontFamily(FONT_NAME);
            customerAppendixHeaderRun.setFontSize(FONT_SIZE);
            setRunTextWithBreaks(customerAppendixHeaderRun,
                    "Приложения к заявке № " + applicationNumber + "\n" + CUSTOMER_APPENDIX_HEADER);

            XWPFParagraph customerAppendixTitle = document.createParagraph();
            customerAppendixTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(customerAppendixTitle);
            XWPFRun customerAppendixTitleRun = customerAppendixTitle.createRun();
            customerAppendixTitleRun.setText(CUSTOMER_SECTION_HEADER);
            customerAppendixTitleRun.setFontFamily(FONT_NAME);
            customerAppendixTitleRun.setFontSize(FONT_SIZE);
            customerAppendixTitleRun.setBold(true);

            for (String line : customerLines) {
                addParagraphWithLineBreaks(document, line);
            }

            if (!roomParamsRows.isEmpty()) {
                addParagraphWithLineBreaks(document, ROOM_PARAMS_START);
                XWPFTable roomParamsTable = document.createTable(roomParamsRows.size(), 4);
                configureTableLayout(roomParamsTable, new int[]{3600, 2600, 2000, 2200});
                for (int rowIndex = 0; rowIndex < roomParamsRows.size(); rowIndex++) {
                    List<String> row = roomParamsRows.get(rowIndex);
                    boolean isHeader = rowIndex == 0;
                    ParagraphAlignment alignment = isHeader ? ParagraphAlignment.CENTER : ParagraphAlignment.LEFT;
                    for (int colIndex = 0; colIndex < 4; colIndex++) {
                        String value = colIndex < row.size() ? row.get(colIndex) : "";
                        setTableCellText(roomParamsTable.getRow(rowIndex).getCell(colIndex), value,
                                PLAN_TABLE_FONT_SIZE, isHeader, alignment);
                    }
                }
            } else {
                for (String line : roomParamsLines) {
                    addParagraphWithLineBreaks(document, line);
                }
            }

            if (!areaBetweenRoomsLine.isBlank()) {
                addParagraphWithLineBreaks(document, areaBetweenRoomsLine);
            }

            for (String requirementLine : measurementRequirementLines) {
                addParagraphWithLineBreaks(document, requirementLine);
            }

            XWPFParagraph spacerBeforeSignature = document.createParagraph();
            setParagraphSpacing(spacerBeforeSignature);

            addParagraphWithLineBreaks(document,
                    "Представитель заказчика _______________________________________________\n" +
                            "                                                 (Должность, ФИО, контактные данные)  ");

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveRequestFormFile(File mapFile) {
        return RequestFormExporter.resolveRequestFormFile(mapFile);
    }

    private static void addRequestSection(XWPFDocument document,
                                          String applicationNumber,
                                          CustomerInfo customerInfo,
                                          String customerInn,
                                          String workDeadline) {
        String deadlineText = workDeadline == null ? "" : workDeadline.trim();
        String innText = customerInn == null ? "" : customerInn.trim();

        XWPFParagraph title = document.createParagraph();
        title.setAlignment(ParagraphAlignment.CENTER);
        setParagraphSpacing(title);
        XWPFRun titleRun = title.createRun();
        titleRun.setText("Заявка № " + applicationNumber);
        titleRun.setFontFamily(FONT_NAME);
        titleRun.setFontSize(FONT_SIZE);
        titleRun.setBold(true);

        XWPFParagraph spacerAfterTitle = document.createParagraph();
        setParagraphSpacing(spacerAfterTitle);

        XWPFParagraph requestText = document.createParagraph();
        setParagraphSpacing(requestText);
        XWPFRun requestRun = requestText.createRun();
        requestRun.setText("Прошу провести работы по проведению измерений в целях:");
        requestRun.setFontFamily(FONT_NAME);
        requestRun.setFontSize(FONT_SIZE);
        requestRun.setBold(true);

        XWPFTable purposeTable = document.createTable(2, 2);
        configureTableLayout(purposeTable, new int[]{6280, 6280});
        setTableCellText(purposeTable.getRow(0).getCell(0), "Поставить отметку", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(purposeTable.getRow(0).getCell(1), "Цель", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(purposeTable.getRow(1).getCell(0), "Χ", FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(purposeTable.getRow(1).getCell(1), "Иная цель:\nВвод здания в эксплуатацию",
                FONT_SIZE, false, ParagraphAlignment.LEFT);

        XWPFParagraph spacerAfterPurpose = document.createParagraph();
        setParagraphSpacing(spacerAfterPurpose);

        XWPFTable optionsTable = document.createTable(16, 3);
        configureTableLayout(optionsTable, new int[]{5652, 5652, 1256});
        setTableCellText(optionsTable.getRow(0).getCell(0), "Протоколы измерений", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(optionsTable.getRow(0).getCell(1), "выдать на руки", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(0).getCell(2), "Х", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(1).getCell(1), "направить электронной почтой", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(2).getCell(1),
                "направить почтовой/курьерской службой (оплата услуг по доставке осуществляется Заказчиком)",
                FONT_SIZE, false, ParagraphAlignment.LEFT);

        setTableCellText(optionsTable.getRow(3).getCell(0),
                "Заявитель оставляет право выбора оптимального метода (методики) измерений по заявке, " +
                        "а также факторов, объектов, точек и сроков исследований (испытаний), измерений",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(3).getCell(1), "ДА", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(4).getCell(0), "Выдать результат:", FONT_SIZE, false,
                ParagraphAlignment.LEFT);

        setTableCellText(optionsTable.getRow(5).getCell(0), "Без указания неопределенности/погрешности",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(5).getCell(1), "на усмотрение ИЛ", FONT_SIZE, false,
                ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(6).getCell(0),
                "В соответствии с показателями качества, установленными в методике (погрешность или " +
                        "неопределенность в зависимости от методики, по которой проводится измерение)",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(6).getCell(1), "на усмотрение ИЛ", FONT_SIZE, false,
                ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(7).getCell(0), "С указанием неопределенности", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(7).getCell(1), "на усмотрение ИЛ", FONT_SIZE, false,
                ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(8).getCell(0), "С указанием погрешности", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(8).getCell(1), "на усмотрение ИЛ", FONT_SIZE, false,
                ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(9).getCell(0),
                "Необходимость указания требований к объекту измерений в протоколе измерений:",
                FONT_SIZE, true, ParagraphAlignment.LEFT);

        setTableCellText(optionsTable.getRow(10).getCell(0),
                "Документы, устанавливающие требования к объекту измерений (при необходимости):",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(10).getCell(1), "ДА", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(11).getCell(0),
                "Особые указания от Заказчика для проведения работ",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(11).getCell(1), "НЕТ", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(12).getCell(0),
                "Заказчик согласен на передачу отчетов во ФГИС Росаккредитация",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(12).getCell(1), "ДА", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(13).getCell(0),
                "Необходимо предоставление доступа для наблюдения за лабораторной деятельностью " +
                        "Заказчику в местах временных работ на объектах Заказчика",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(13).getCell(1), "ДА", FONT_SIZE, false, ParagraphAlignment.CENTER);

        setTableCellText(optionsTable.getRow(14).getCell(0),
                "Данные, предоставленные Заказчиком, за которые он несет ответственность",
                FONT_SIZE, false, ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(14).getCell(1), "Приложение к заявке", FONT_SIZE, false,
                ParagraphAlignment.LEFT);

        setTableCellText(optionsTable.getRow(15).getCell(0), "Сроки выполнения работ", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(optionsTable.getRow(15).getCell(1), deadlineText, FONT_SIZE, false,
                ParagraphAlignment.LEFT);

        mergeCellsVertically(optionsTable, 0, 0, 2);
        mergeCellsHorizontally(optionsTable, 3, 1, 2);
        mergeCellsHorizontally(optionsTable, 4, 0, 2);
        mergeCellsHorizontally(optionsTable, 5, 1, 2);
        mergeCellsHorizontally(optionsTable, 6, 1, 2);
        mergeCellsHorizontally(optionsTable, 7, 1, 2);
        mergeCellsHorizontally(optionsTable, 8, 1, 2);
        mergeCellsHorizontally(optionsTable, 9, 0, 2);
        mergeCellsHorizontally(optionsTable, 10, 1, 2);
        mergeCellsHorizontally(optionsTable, 11, 1, 2);
        mergeCellsHorizontally(optionsTable, 12, 1, 2);
        mergeCellsHorizontally(optionsTable, 13, 1, 2);
        mergeCellsHorizontally(optionsTable, 14, 1, 2);
        mergeCellsHorizontally(optionsTable, 15, 1, 2);

        XWPFParagraph spacerBeforeCustomer = document.createParagraph();
        setParagraphSpacing(spacerBeforeCustomer);

        XWPFParagraph customerTitle = document.createParagraph();
        setParagraphSpacing(customerTitle);
        XWPFRun customerRun = customerTitle.createRun();
        customerRun.setText("Сведения о заказчике:");
        customerRun.setFontFamily(FONT_NAME);
        customerRun.setFontSize(FONT_SIZE);

        XWPFTable customerTable = document.createTable(4, 2);
        configureTableLayout(customerTable, new int[]{3770, 8790});
        setTableCellText(customerTable.getRow(0).getCell(0), "Наименование", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(0).getCell(1), customerInfo.name, FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(1).getCell(0), "ИНН", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(1).getCell(1), innText, FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(2).getCell(0), "Электронная почта", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(2).getCell(1), customerInfo.email, FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(3).getCell(0), "Контактный телефон", FONT_SIZE, false,
                ParagraphAlignment.LEFT);
        setTableCellText(customerTable.getRow(3).getCell(1), customerInfo.phone, FONT_SIZE, false,
                ParagraphAlignment.LEFT);

        XWPFParagraph spacerAfterCustomer = document.createParagraph();
        setParagraphSpacing(spacerAfterCustomer);

        addParagraphWithLineBreaks(document,
                "Заявитель ознакомлен с методами и методиками, используемыми испытательной лабораторией.\n\n" +
                        "Заказчик согласен на отклонение от методов (методик) измерений с учетом выполненного " +
                        "ИЛ технического обоснования отклонений, если такое отклонение потребуется при " +
                        "выполнении заявки.\n\n" +
                        "Заявитель ознакомлен с тем, что ИЛ не использует в своей работе методов, не изложенных " +
                        "в нормативных документах (домашних методов): нестандартных методик; методик, " +
                        "разработанных ИЛ; стандартных методик, используемых за пределами целевой области их " +
                        "применений; расширений и модификаций стандартных методик.\n\n" +
                        "Заявитель обязуется:\n" +
                        "- обеспечить доступ на объект для проведения исследований, испытаний, измерений;\n" +
                        "- предоставить всю необходимую информацию для проведения работ по заявке.\n\n" +
                        "Заявитель: ");

        XWPFParagraph spacerBeforeSignature = document.createParagraph();
        setParagraphSpacing(spacerBeforeSignature);

        XWPFTable signatureTable = document.createTable(2, 3);
        configureTableLayout(signatureTable, new int[]{4187, 4187, 4186});
        setTableCellText(signatureTable.getRow(0).getCell(0), "______________________", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(signatureTable.getRow(0).getCell(1), "______________________", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(signatureTable.getRow(0).getCell(2), "______________________", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(signatureTable.getRow(1).getCell(0), "должность", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(signatureTable.getRow(1).getCell(1), "подпись", FONT_SIZE, false,
                ParagraphAlignment.CENTER);
        setTableCellText(signatureTable.getRow(1).getCell(2), "инициалы, фамилия", FONT_SIZE, false,
                ParagraphAlignment.CENTER);

        addPageBreak(document);
    }

    private static String resolveApplicationNumber(File protocolFile, File mapFile) {
        String fromProtocol = extractApplicationNumberFromProtocol(protocolFile);
        if (!fromProtocol.isBlank()) {
            return fromProtocol;
        }
        return RequestFormExporter.resolveApplicationNumberFromMap(mapFile);
    }

    private static String extractApplicationNumberFromProtocol(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return "";
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document, true);
            return extractApplicationNumberFromLines(lines);
        } catch (Exception ignored) {
            return "";
        }
    }

    private static String extractApplicationNumberFromLines(List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String basisLabel = APPLICATION_BASIS_LABEL.toLowerCase(Locale.ROOT);
        String basisAltLabel = APPLICATION_BASIS_ALT_LABEL.toLowerCase(Locale.ROOT);
        for (String line : lines) {
            String normalized = normalizeSpace(line);
            String lower = normalized.toLowerCase(Locale.ROOT);
            if (!lower.contains(basisLabel) && !lower.contains(basisAltLabel)) {
                continue;
            }
            int applicationIndex = lower.indexOf("заявка");
            if (applicationIndex < 0) {
                continue;
            }
            String tail = normalized.substring(applicationIndex + "заявка".length()).trim();
            return trimLeadingPunctuation(tail);
        }
        return "";
    }

    private static List<List<String>> extractPlanRows(File planFile) {
        List<List<String>> rows = new ArrayList<>();
        if (planFile == null || !planFile.exists()) {
            return rows;
        }
        try (InputStream inputStream = new FileInputStream(planFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            for (XWPFTable table : document.getTables()) {
                if (!isPlanTable(table)) {
                    continue;
                }
                for (XWPFTableRow row : table.getRows()) {
                    List<String> values = new ArrayList<>();
                    values.add(getCellText(row, 0));
                    values.add(getCellText(row, 1));
                    values.add(getCellText(row, 3));
                    values.add(getCellText(row, 4));
                    values.add(getCellText(row, 5));
                    rows.add(values);
                }
                break;
            }
        } catch (Exception ignored) {
            // пропускаем чтение плана измерений при ошибке
        }
        return rows;
    }

    private static void applyPlanDeadline(List<List<String>> rows, String workDeadline) {
        if (rows == null || rows.size() < 2) {
            return;
        }
        if (workDeadline == null || workDeadline.isBlank()) {
            return;
        }
        List<String> row = rows.get(1);
        while (row.size() < 5) {
            row.add("");
        }
        row.set(4, workDeadline.trim());
    }

    private static boolean isPlanTable(XWPFTable table) {
        if (table == null || table.getNumberOfRows() == 0) {
            return false;
        }
        XWPFTableRow row = table.getRow(0);
        if (row == null || row.getTableCells().size() < 6) {
            return false;
        }
        String header = normalizeSpace(getCellText(row, 0));
        return header.toLowerCase(Locale.ROOT).contains(PLAN_TABLE_HEADER.toLowerCase(Locale.ROOT));
    }

    private static String getCellText(XWPFTableRow row, int index) {
        if (row == null || row.getTableCells().size() <= index) {
            return "";
        }
        XWPFTableCell cell = row.getCell(index);
        return cell == null ? "" : normalizeSpace(cell.getText());
    }

    private static CustomerInfo extractCustomerInfo(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return new CustomerInfo("", "", "");
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            String value = extractCustomerInfoValue(extractLines(document, true));
            return parseCustomerInfo(value);
        } catch (Exception ignored) {
            return new CustomerInfo("", "", "");
        }
    }

    private static String extractCustomerInfoValue(List<String> lines) {
        if (lines == null || lines.isEmpty()) {
            return "";
        }
        String label = CUSTOMER_INFO_LABEL.toLowerCase(Locale.ROOT);
        for (int index = 0; index < lines.size(); index++) {
            String line = normalizeSpace(lines.get(index));
            String lower = line.toLowerCase(Locale.ROOT);
            int labelIndex = lower.indexOf(label);
            if (labelIndex < 0) {
                continue;
            }
            String tail = trimLeadingPunctuation(line.substring(labelIndex + CUSTOMER_INFO_LABEL.length()));
            if (!tail.isBlank()) {
                return tail;
            }
            List<String> values = new ArrayList<>();
            for (int nextIndex = index + 1; nextIndex < lines.size(); nextIndex++) {
                String next = normalizeSpace(lines.get(nextIndex));
                if (next.matches("^\\d+\\.\\s+.*")) {
                    break;
                }
                if (!next.isBlank()) {
                    values.add(next);
                }
            }
            return normalizeSpace(String.join(" ", values));
        }
        return "";
    }

    private static CustomerInfo parseCustomerInfo(String value) {
        String normalized = normalizeSpace(value);
        if (normalized.isBlank()) {
            return new CustomerInfo("", "", "");
        }
        Matcher emailMatcher = EMAIL_PATTERN.matcher(normalized);
        if (emailMatcher.find()) {
            String name = stripCustomerSeparators(normalized.substring(0, emailMatcher.start()));
            String email = emailMatcher.group();
            String phone = cleanCustomerPhone(normalized.substring(emailMatcher.end()));
            return new CustomerInfo(name, email, phone);
        }
        List<String> parts = splitCustomerParts(normalized);
        String name = parts.isEmpty() ? normalized : parts.get(0);
        String email = parts.size() > 1 ? parts.get(1) : "";
        String phone = parts.size() > 2 ? cleanCustomerPhone(parts.get(2)) : "";
        return new CustomerInfo(name, email, phone);
    }

    private static List<String> splitCustomerParts(String value) {
        List<String> result = new ArrayList<>();
        for (String part : value.split("\\.")) {
            String trimmed = stripCustomerSeparators(part);
            if (!trimmed.isBlank()) {
                result.add(trimmed);
            }
        }
        return result;
    }

    private static String cleanCustomerPhone(String value) {
        String result = stripCustomerSeparators(value);
        result = PHONE_PREFIX_PATTERN.matcher(result).replaceFirst("");
        return stripCustomerSeparators(result);
    }

    private static String stripCustomerSeparators(String value) {
        return normalizeSpace(value)
                .replaceFirst("^[\\s,.;:]+", "")
                .replaceFirst("[\\s,.;:]+$", "")
                .trim();
    }

    private static List<String> extractCustomerDataLines(File protocolFile) {
        List<String> lines = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return lines;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            lines.addAll(extractSection(allLines, CUSTOMER_DATA_START, CUSTOMER_DATA_END));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return lines;
    }

    private static List<String> extractRoomParamsLines(File protocolFile) {
        List<String> lines = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return lines;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            lines.addAll(extractSection(allLines, ROOM_PARAMS_START, ROOM_PARAMS_END));
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return lines;
    }

    private static String extractAreaBetweenRoomsLine(File protocolFile) {
        if (protocolFile == null || !protocolFile.exists()) {
            return "";
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> allLines = extractLines(document, true);
            for (String line : allLines) {
                String normalized = normalizeSpace(line);
                if (normalized.toLowerCase(Locale.ROOT)
                        .contains(AREA_BETWEEN_ROOMS_MARKER.toLowerCase(Locale.ROOT))) {
                    return normalized;
                }
            }
        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }
        return "";
    }

    private static List<List<String>> extractRoomParamsTable(File protocolFile) {
        List<List<String>> rows = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return rows;
        }
        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            XWPFTable table = findTableAfterTitle(document, ROOM_PARAMS_START);
            if (table == null) {
                table = findTableAfterTitle(document, ROOM_PARAMS_ALT_START);
            }
            if (table == null) {
                return rows;
            }
            for (XWPFTableRow row : table.getRows()) {
                List<String> values = new ArrayList<>();
                for (int colIndex = 0; colIndex < 4; colIndex++) {
                    values.add(getCellText(row, colIndex));
                }
                rows.add(values);
            }
        } catch (Exception ignored) {
            // пропускаем извлечение таблицы при ошибке
        }
        return rows;
    }

    private static List<String> extractSection(List<String> lines, String startMarker, String endMarker) {
        List<String> result = new ArrayList<>();
        if (lines == null || lines.isEmpty()) {
            return result;
        }
        String lowerStart = startMarker.toLowerCase(Locale.ROOT);
        String lowerEnd = endMarker.toLowerCase(Locale.ROOT);
        boolean collecting = false;
        for (String line : lines) {
            String normalized = normalizeSpace(line);
            String lower = normalized.toLowerCase(Locale.ROOT);
            if (!collecting) {
                int startIndex = lower.indexOf(lowerStart);
                if (startIndex < 0) {
                    continue;
                }
                String value = normalized.substring(startIndex);
                collecting = true;
                if (appendUntilEnd(result, value, lowerEnd)) {
                    break;
                }
                continue;
            }
            if (appendUntilEnd(result, normalized, lowerEnd)) {
                break;
            }
        }
        return sanitizeCustomerLines(result);
    }

    private static boolean appendUntilEnd(List<String> result, String value, String lowerEndMarker) {
        if (value == null) {
            return false;
        }
        String lower = value.toLowerCase(Locale.ROOT);
        int endIndex = lower.indexOf(lowerEndMarker);
        if (endIndex >= 0) {
            String fragment = value.substring(0, endIndex).trim();
            if (!fragment.isBlank()) {
                result.add(fragment);
            }
            return true;
        }
        if (!value.isBlank()) {
            result.add(value);
        }
        return false;
    }

    private static List<String> sanitizeCustomerLines(List<String> lines) {
        List<String> sanitized = new ArrayList<>();
        for (String line : lines) {
            if (line == null || line.isBlank()) {
                continue;
            }
            String value = line;
            if (value.toLowerCase(Locale.ROOT).contains("13.1 конструктивные решения")) {
                value = value.replaceFirst("(?i)^13\\.1\\s+", "");
            }
            sanitized.add(value);
        }
        return sanitized;
    }

    private static List<String> extractLines(XWPFDocument document, boolean includeTables) {
        List<String> lines = new ArrayList<>();
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                addParagraphLines(lines, paragraph.getText());
            } else if (includeTables && element instanceof XWPFTable table) {
                addTableLines(lines, table);
            }
        }
        return lines;
    }

    private static void addParagraphLines(List<String> lines, String text) {
        if (text == null) {
            return;
        }
        String[] split = text.split("\\R");
        for (String line : split) {
            String normalized = normalizeSpace(line);
            if (!normalized.isBlank()) {
                lines.add(normalized);
            }
        }
    }

    private static void addTableLines(List<String> lines, XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            StringBuilder builder = new StringBuilder();
            for (XWPFTableCell cell : row.getTableCells()) {
                String cellText = normalizeSpace(cell.getText());
                if (cellText.isBlank()) {
                    continue;
                }
                if (builder.length() > 0) {
                    builder.append(' ');
                }
                builder.append(cellText);
            }
            String value = builder.toString().trim();
            if (!value.isBlank()) {
                lines.add(value);
            }
        }
    }

    private static void setParagraphSpacing(XWPFParagraph paragraph) {
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
    }

    private static void addParagraphWithLineBreaks(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
        setRunTextWithBreaks(run, text);
    }

    private static void addPageBreak(XWPFDocument document) {
        XWPFParagraph pageBreakParagraph = document.createParagraph();
        setParagraphSpacing(pageBreakParagraph);
        pageBreakParagraph.setPageBreak(true);
    }

    private static List<String> extractMeasurementRequirementLines(File protocolFile) {
        List<String> result = new ArrayList<>();
        if (protocolFile == null || !protocolFile.exists()) {
            return result;
        }
        if (!protocolFile.getName().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            return result;
        }

        try (InputStream inputStream = new FileInputStream(protocolFile);
             XWPFDocument document = new XWPFDocument(inputStream)) {

            List<String> lnRequirements = extractLnRequirements(document);

            List<String> rwFloorRequirements = extractRRequirements(
                    document,
                    RW_FLOOR_SENTENCE_MARKER,
                    "Для перекрытия между помещениями "
            );

            List<String> rwPartitionRequirements = extractRRequirements(
                    document,
                    RW_PARTITION_SENTENCE_MARKER,
                    "Для перегородки между помещениями "
            );

            appendGroup(result, lnRequirements);
            appendGroup(result, rwFloorRequirements);
            appendGroup(result, rwPartitionRequirements);

        } catch (Exception ignored) {
            // пропускаем извлечение данных при ошибке
        }

        return result;
    }

    private static void appendGroup(List<String> target, List<String> group) {
        if (group.isEmpty()) {
            return;
        }
        if (!target.isEmpty()) {
            target.add("");
        }
        target.addAll(group);
    }

    private static List<String> extractLnRequirements(XWPFDocument document) {
        List<String> result = new ArrayList<>();
        if (document == null) {
            return result;
        }

        List<IBodyElement> elements = document.getBodyElements();

        int start = findSectionStartIndex(elements, "17.1", 0);
        if (start < 0) {
            return result;
        }

        int end = findSectionStartIndex(elements, "17.2", start + 1);
        if (end < 0) {
            end = elements.size();
        }

        // Логика как ты описал: в 17.1 ищем таблицу где есть "Ln, дБ",
        // затем берём абзац сразу после этой таблицы и парсим в нём LNW_SENTENCE_MARKER + (нормативные требования ...)
        for (int i = start; i < end; i++) {
            IBodyElement element = elements.get(i);
            if (!(element instanceof XWPFTable table)) {
                continue;
            }
            if (!tableContainsLnDbMarker(table)) {
                continue;
            }

            String nextParagraphText = findNextNonEmptyParagraphText(elements, i + 1, end);
            if (nextParagraphText.isBlank()) {
                continue;
            }

            result.addAll(extractRequirementLinesFromText(
                    nextParagraphText,
                    LNW_SENTENCE_MARKER,
                    "Для перекрытия между помещениями "
            ));
        }

        return distinctPreserveOrder(result);
    }


    private static List<String> extractRRequirements(XWPFDocument document, String sentenceMarker, String prefix) {
        List<String> result = new ArrayList<>();
        if (document == null || sentenceMarker == null || sentenceMarker.isBlank()) {
            return result;
        }

        List<IBodyElement> elements = document.getBodyElements();

        int start = findSectionStartIndex(elements, "17.2", 0);
        if (start < 0) {
            return result;
        }

        // Важно: конец секции ищем как ЗАГОЛОВОК "18." (в начале абзаца), а не просто contains("18.")
        int end = findSectionStartIndex(elements, "18.", start + 1);
        if (end < 0) {
            end = elements.size();
        }

        for (int i = start; i < end; i++) {
            IBodyElement element = elements.get(i);

            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (text.isBlank()) {
                    continue;
                }
                result.addAll(extractRequirementLinesFromText(text, sentenceMarker, prefix));
                continue;
            }

            // На всякий случай: если фраза почему-то попала в таблицу
            if (element instanceof XWPFTable table) {
                for (XWPFTableRow row : table.getRows()) {
                    String rowText = tableRowToText(row);
                    if (rowText.isBlank()) {
                        continue;
                    }
                    result.addAll(extractRequirementLinesFromText(rowText, sentenceMarker, prefix));
                }
            }
        }

        return distinctPreserveOrder(result);
    }

    private static int findLineIndexContaining(List<String> lines, String marker) {
        if (lines == null || marker == null || marker.isBlank()) {
            return -1;
        }
        for (int i = 0; i < lines.size(); i++) {
            if (containsIgnoreCase(lines.get(i), marker)) {
                return i;
            }
        }
        return -1;
    }

    private static String buildSearchWindow(List<String> lines, int start, int end) {
        StringBuilder builder = new StringBuilder();
        int max = Math.min(end, start + 24);
        for (int i = start; i < max; i++) {
            String value = normalizeSpace(lines.get(i));
            if (value.matches("^\\d+\\.\\d+.*") || value.matches("^\\d+\\..*")) {
                break;
            }
            if (builder.length() > 0) {
                builder.append(' ');
            }
            builder.append(value);
        }
        return builder.toString();
    }

    private static String extractBetweenMarkers(String source, String startMarker, String endMarker) {
        if (source == null || source.isBlank()) {
            return "";
        }
        String lower = source.toLowerCase(Locale.ROOT);
        String lowerStart = startMarker.toLowerCase(Locale.ROOT);
        String lowerEnd = endMarker.toLowerCase(Locale.ROOT);
        int startIndex = lower.indexOf(lowerStart);
        if (startIndex < 0) {
            return "";
        }
        int fromIndex = startIndex + lowerStart.length();
        int endIndex = lower.indexOf(lowerEnd, fromIndex);
        if (endIndex < 0) {
            return "";
        }
        return normalizeSpace(source.substring(fromIndex, endIndex));
    }

    private static String extractNormativeRequirements(String source) {
        if (source == null || source.isBlank()) {
            return "";
        }
        Pattern pattern = Pattern.compile("\\(\\s*нормативные требования\\s*([^)]*)\\)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);
        Matcher matcher = pattern.matcher(source);
        if (!matcher.find()) {
            return "";
        }
        return normalizeSpace(matcher.group(1));
    }

    private static boolean containsIgnoreCase(String value, String marker) {
        if (value == null || marker == null) {
            return false;
        }
        return normalizeSpace(value).toLowerCase(Locale.ROOT)
                .contains(normalizeSpace(marker).toLowerCase(Locale.ROOT));
    }

    private static void setRunTextWithBreaks(XWPFRun run, String text) {
        if (text == null) {
            return;
        }
        String[] parts = text.split("\\n", -1);
        int textPos = 0;
        for (int index = 0; index < parts.length; index++) {
            if (index > 0) {
                run.addBreak();
            }
            run.setText(parts[index], textPos++);
        }
    }

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(scaleWidth(12560)));

        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable.CENTER);

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
        int[] scaledWidths = scaleColumnWidths(columnWidths);
        for (int width : scaledWidths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }

        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < scaledWidths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, scaledWidths[colIndex]);
            }
        }
    }

    private static int[] scaleColumnWidths(int[] columnWidths) {
        int[] scaled = new int[columnWidths.length];
        for (int i = 0; i < columnWidths.length; i++) {
            scaled[i] = scaleWidth(columnWidths[i]);
        }
        return scaled;
    }

    private static int scaleWidth(int width) {
        return Math.max(1, (int) Math.round(width * REQUEST_TABLE_WIDTH_SCALE));
    }

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthDxa));
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
        if (fromCol >= toCol) {
            return;
        }
        XWPFTableRow tableRow = table.getRow(row);
        if (tableRow == null) {
            return;
        }
        XWPFTableCell cell = tableRow.getCell(fromCol);
        if (cell == null) {
            return;
        }
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        tcPr.addNewGridSpan().setVal(BigInteger.valueOf(toCol - fromCol + 1));

        for (int colIndex = toCol; colIndex > fromCol; colIndex--) {
            tableRow.removeCell(colIndex);
        }
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize, boolean bold,
                                         ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        run.setText(text);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);
        run.setBold(bold);
    }

    private static String normalizeSpace(String value) {
        if (value == null) {
            return "";
        }
        return value.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }

    private static String trimLeadingPunctuation(String value) {
        if (value == null) {
            return "";
        }
        int index = 0;
        while (index < value.length() && !Character.isLetterOrDigit(value.charAt(index))) {
            index++;
        }
        return value.substring(index).trim();
    }

    private static XWPFTable findTableAfterTitle(XWPFDocument document, String title) {
        if (document == null) {
            return null;
        }
        String lowerTitle = normalizeSpace(title).toLowerCase(Locale.ROOT);
        List<IBodyElement> elements = document.getBodyElements();
        for (int i = 0; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText()).toLowerCase(Locale.ROOT);
                if (!text.contains(lowerTitle)) {
                    continue;
                }
                for (int j = i + 1; j < elements.size(); j++) {
                    IBodyElement next = elements.get(j);
                    if (next instanceof XWPFTable table) {
                        return table;
                    }
                }
            } else if (element instanceof XWPFTable table) {
                if (tableContainsTitle(table, lowerTitle)) {
                    return table;
                }
            }
        }
        return null;
    }

    private static boolean tableContainsTitle(XWPFTable table, String lowerTitle) {
        if (table == null) {
            return false;
        }
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                String text = normalizeSpace(cell.getText()).toLowerCase(Locale.ROOT);
                if (text.contains(lowerTitle)) {
                    return true;
                }
            }
        }
        return false;
    }
    private static int findSectionStartIndex(List<IBodyElement> elements, String sectionNumber, int fromIndex) {
        if (elements == null || elements.isEmpty() || sectionNumber == null || sectionNumber.isBlank()) {
            return -1;
        }

        // Ищем заголовок вида "17.1." / "17.1 " / "18." и т.п.
        // Важно: НЕ должны матчиться подзаголовки типа "17.1.1"
        String base = normalizeSpace(sectionNumber);
        if (base.endsWith(".")) {
            base = base.substring(0, base.length() - 1).trim();
        }
        if (base.isBlank()) {
            return -1;
        }

        Pattern headingPattern = Pattern.compile(
                "^\\s*" + Pattern.quote(base) + "\\s*\\.?\\s*(?!\\d)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
        );

        int start = Math.max(0, fromIndex);
        for (int i = start; i < elements.size(); i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (headingPattern.matcher(text).find()) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static String findNextNonEmptyParagraphText(List<IBodyElement> elements, int fromIndex, int endExclusive) {
        if (elements == null || elements.isEmpty()) {
            return "";
        }
        int to = Math.min(endExclusive, elements.size());
        for (int i = Math.max(0, fromIndex); i < to; i++) {
            IBodyElement element = elements.get(i);
            if (element instanceof XWPFParagraph paragraph) {
                String text = normalizeSpace(paragraph.getText());
                if (!text.isBlank()) {
                    return text;
                }
            }
        }
        return "";
    }

    private static boolean tableContainsLnDbMarker(XWPFTable table) {
        if (table == null) {
            return false;
        }
        Pattern lnPattern = Pattern.compile("Ln\\s*,\\s*дБ", Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE);

        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                String text = normalizeSpace(cell.getText());
                if (lnPattern.matcher(text).find()) {
                    return true;
                }
            }
        }
        return false;
    }

    private static String tableRowToText(XWPFTableRow row) {
        if (row == null) {
            return "";
        }
        StringBuilder builder = new StringBuilder();
        for (XWPFTableCell cell : row.getTableCells()) {
            String cellText = normalizeSpace(cell.getText());
            if (cellText.isBlank()) {
                continue;
            }
            if (builder.length() > 0) {
                builder.append(' ');
            }
            builder.append(cellText);
        }
        return builder.toString().trim();
    }

    private static List<String> extractRequirementLinesFromText(String text, String sentenceMarker, String prefix) {
        List<String> result = new ArrayList<>();
        if (text == null || text.isBlank() || sentenceMarker == null || sentenceMarker.isBlank()) {
            return result;
        }

        String normalized = normalizeSpace(text);

        // Вытаскиваем строго "A и B" после маркера и норму из "(нормативные требования ...)"
        Pattern pattern = Pattern.compile(
                Pattern.quote(sentenceMarker)
                        + "\\s*([A-Za-zА-Яа-я0-9]+)\\s+и\\s+([A-Za-zА-Яа-я0-9]+).*?"
                        + "\\(\\s*нормативные требования\\s*([^)]*)\\)",
                Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE
        );

        Matcher matcher = pattern.matcher(normalized);
        while (matcher.find()) {
            String first = normalizeSpace(matcher.group(1));
            String second = normalizeSpace(matcher.group(2));
            String norm = normalizeSpace(matcher.group(3));

            if (first.isBlank() || second.isBlank() || norm.isBlank()) {
                continue;
            }

            String rooms = first + " и " + second;
            String line = (prefix == null ? "" : prefix) + rooms + " (нормативные требования " + norm + ")";
            result.add(normalizeSpace(line));
        }

        return result;
    }

    private static List<String> distinctPreserveOrder(List<String> values) {
        java.util.LinkedHashSet<String> set = new java.util.LinkedHashSet<>();
        if (values != null) {
            for (String value : values) {
                String v = normalizeSpace(value);
                if (!v.isBlank()) {
                    set.add(v);
                }
            }
        }
        return new ArrayList<>(set);
    }

    private record CustomerInfo(String name, String email, String phone) {
    }

}
