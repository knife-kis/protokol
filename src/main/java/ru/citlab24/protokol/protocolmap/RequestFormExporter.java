package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
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
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

final class RequestFormExporter {
    private static final String REQUEST_FORM_NAME = "заявка.docx";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int PLAN_TABLE_FONT_SIZE = 10;
    private static final int MICROCLIMATE_HEADER_MAIN_FONT_SIZE = 10;
    private static final int MICROCLIMATE_TABLE_FONT_SIZE = 9;
    private static final int VENTILATION_TABLE_FONT_SIZE = 9;
    private static final int MED_TABLE_FONT_SIZE = 9;
    private static final int VENTILATION_SOURCE_START_ROW = 4;
    private static final int MED_SOURCE_START_ROW = 5;
    private static final int MAP_APPLICATION_ROW_INDEX = 22;
    private static final int MAP_CUSTOMER_ROW_INDEX = 5;
    private static final double REQUEST_TABLE_WIDTH_SCALE = 0.8;
    private static final String OBJECT_PREFIX = "4. Наименование объекта:";
    private static final String ADDRESS_PREFIX = "Адрес объекта";
    private static final String NORMATIVE_SECTION_TITLE = "Сведения о нормативных документах";
    private static final String NORMATIVE_HEADER_TITLE = "Измеряемый показатель";

    private RequestFormExporter() {
    }

    static void generate(File sourceFile, File mapFile, String workDeadline, String customerInn) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveRequestFormFile(mapFile);
        String applicationNumber = resolveApplicationNumberFromMap(mapFile);
        CustomerInfo customerInfo = resolveCustomerInfo(mapFile);
        String objectName = resolveObjectName(mapFile);
        String objectAddress = resolveObjectAddress(mapFile);
        List<NormativeRow> normativeRows = resolveNormativeRows(sourceFile);
        List<MicroclimateRow> microclimateRows = resolveMicroclimateRows(sourceFile);
        List<VentilationRow> ventilationRows = resolveVentilationRows(sourceFile);
        String ventilationMethod = resolveVentilationNormativeMethod(sourceFile);
        List<String> medRows = resolveMedRows(sourceFile);
        String medNormativeMethod = resolveMedNormativeMethod(sourceFile);
        boolean hasMedSheet = hasSheetByName(sourceFile, "МЭД");
        if (normativeRows.isEmpty()) {
            normativeRows.add(new NormativeRow("", ""));
        }
        String protocolDate = resolveProtocolDateFromSource(sourceFile);
        String deadlineText = workDeadline == null ? "" : workDeadline.trim();
        String planDeadline = deadlineText.isEmpty() ? protocolDate : deadlineText;
        String innText = customerInn == null ? "" : customerInn.trim();

        try (XWPFDocument document = new XWPFDocument()) {
            applyStandardHeader(document);

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
            setTableCellText(purposeTable.getRow(0).getCell(0), "Поставить отметку", false, ParagraphAlignment.LEFT);
            setTableCellText(purposeTable.getRow(0).getCell(1), "Цель", false, ParagraphAlignment.LEFT);
            setTableCellText(purposeTable.getRow(1).getCell(0), "Χ", false, ParagraphAlignment.LEFT);
            setTableCellText(purposeTable.getRow(1).getCell(1),
                    "Иная цель:\nВвод здания в эксплуатацию", false, ParagraphAlignment.LEFT);

            XWPFParagraph spacerAfterPurpose = document.createParagraph();
            setParagraphSpacing(spacerAfterPurpose);

            XWPFTable optionsTable = document.createTable(16, 3);
            configureTableLayout(optionsTable, new int[]{5652, 5652, 1256});
            setTableCellText(optionsTable.getRow(0).getCell(0), "Протоколы измерений", false,
                    ParagraphAlignment.CENTER);
            setTableCellText(optionsTable.getRow(0).getCell(1), "выдать на руки", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(0).getCell(2), "Х", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(1).getCell(1), "направить электронной почтой", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(2).getCell(1),
                    "направить почтовой/курьерской службой (оплата услуг по доставке осуществляется Заказчиком)",
                    false, ParagraphAlignment.LEFT);

            setTableCellText(optionsTable.getRow(3).getCell(0),
                    "Заявитель оставляет право выбора оптимального метода (методики) измерений по заявке, " +
                            "а также факторов, объектов, точек и сроков исследований (испытаний), измерений",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(3).getCell(1), "ДА", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(4).getCell(0), "Выдать результат:", false,
                    ParagraphAlignment.LEFT);

            setTableCellText(optionsTable.getRow(5).getCell(0), "Без указания неопределенности/погрешности",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(5).getCell(1), "на усмотрение ИЛ", false,
                    ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(6).getCell(0),
                    "В соответствии с показателями качества, установленными в методике (погрешность или " +
                            "неопределенность в зависимости от методики, по которой проводится измерение)",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(6).getCell(1), "на усмотрение ИЛ", false,
                    ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(7).getCell(0), "С указанием неопределенности", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(7).getCell(1), "на усмотрение ИЛ", false,
                    ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(8).getCell(0), "С указанием погрешности", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(8).getCell(1), "на усмотрение ИЛ", false,
                    ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(9).getCell(0),
                    "Необходимость указания требований к объекту измерений в протоколе измерений:",
                    true, ParagraphAlignment.LEFT);

            setTableCellText(optionsTable.getRow(10).getCell(0),
                    "Документы, устанавливающие требования к объекту измерений (при необходимости):",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(10).getCell(1), "ДА", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(11).getCell(0),
                    "Особые указания от Заказчика для проведения работ",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(11).getCell(1), "НЕТ", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(12).getCell(0),
                    "Заказчик согласен на передачу отчетов во ФГИС Росаккредитация",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(12).getCell(1), "ДА", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(13).getCell(0),
                    "Необходимо предоставление доступа для наблюдения за лабораторной деятельностью " +
                            "Заказчику в местах временных работ на объектах Заказчика",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(13).getCell(1), "ДА", false, ParagraphAlignment.CENTER);

            setTableCellText(optionsTable.getRow(14).getCell(0),
                    "Данные, предоставленные Заказчиком, за которые он несет ответственность",
                    false, ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(14).getCell(1), "Приложение к заявке", false,
                    ParagraphAlignment.LEFT);

            setTableCellText(optionsTable.getRow(15).getCell(0), "Сроки выполнения работ", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(optionsTable.getRow(15).getCell(1), deadlineText, false,
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
            setTableCellText(customerTable.getRow(0).getCell(0), "Наименование", false, ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(0).getCell(1), customerInfo.name, false, ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(1).getCell(0), "ИНН", false, ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(1).getCell(1), innText, false, ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(2).getCell(0), "Электронная почта", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(2).getCell(1), customerInfo.email, false,
                    ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(3).getCell(0), "Контактный телефон", false,
                    ParagraphAlignment.LEFT);
            setTableCellText(customerTable.getRow(3).getCell(1), customerInfo.phone, false,
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
            setTableCellText(signatureTable.getRow(0).getCell(0), "______________________", false,
                    ParagraphAlignment.CENTER);
            setTableCellText(signatureTable.getRow(0).getCell(1), "______________________", false,
                    ParagraphAlignment.CENTER);
            setTableCellText(signatureTable.getRow(0).getCell(2), "______________________", false,
                    ParagraphAlignment.CENTER);

            setTableCellText(signatureTable.getRow(1).getCell(0), "должность", false,
                    ParagraphAlignment.CENTER);
            setTableCellText(signatureTable.getRow(1).getCell(1), "подпись", false,
                    ParagraphAlignment.CENTER);
            setTableCellText(signatureTable.getRow(1).getCell(2), "инициалы, фамилия", false,
                    ParagraphAlignment.CENTER);

            XWPFParagraph pageBreak = document.createParagraph();
            setParagraphSpacing(pageBreak);
            pageBreak.createRun().addBreak(BreakType.PAGE);

            XWPFParagraph appendixHeader = document.createParagraph();
            appendixHeader.setAlignment(ParagraphAlignment.RIGHT);
            setParagraphSpacing(appendixHeader);
            XWPFRun appendixHeaderRun = appendixHeader.createRun();
            appendixHeaderRun.setFontFamily(FONT_NAME);
            appendixHeaderRun.setFontSize(FONT_SIZE);
            setRunTextWithBreaks(appendixHeaderRun,
                    "Приложения к заявке № " + applicationNumber + "\nПриложение – План измерений");

            XWPFParagraph appendixTitle = document.createParagraph();
            appendixTitle.setAlignment(ParagraphAlignment.CENTER);
            setParagraphSpacing(appendixTitle);
            XWPFRun appendixTitleRun = appendixTitle.createRun();
            appendixTitleRun.setText("ПЛАН ИЗМЕРЕНИЙ");
            appendixTitleRun.setFontFamily(FONT_NAME);
            appendixTitleRun.setFontSize(FONT_SIZE);
            appendixTitleRun.setBold(true);

            int dataRowsCount = normativeRows.size();
            XWPFTable planTable = document.createTable(1 + dataRowsCount, 5);
            configureTableLayout(planTable, new int[]{2500, 2500, 2100, 2560, 2900});
            setTableCellText(planTable.getRow(0).getCell(0), "Наименование объекта измерений",
                    PLAN_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(1), "Адрес объекта измерений",
                    PLAN_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(2), "Показатель",
                    PLAN_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(3), "Метод (методика) измерений (шифр)",
                    PLAN_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
            setTableCellText(planTable.getRow(0).getCell(4),
                    "Срок выполнения работ в области лабораторной деятельности и оформления проекта отчета",
                    PLAN_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

            for (int index = 0; index < dataRowsCount; index++) {
                int rowIndex = index + 1;
                NormativeRow row = normativeRows.get(index);
                setTableCellText(planTable.getRow(rowIndex).getCell(2), row.indicator,
                        PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                setTableCellText(planTable.getRow(rowIndex).getCell(3), row.method,
                        PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                if (index == 0) {
                    setTableCellText(planTable.getRow(rowIndex).getCell(0), objectName,
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                    setTableCellText(planTable.getRow(rowIndex).getCell(1), objectAddress,
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                    setTableCellText(planTable.getRow(rowIndex).getCell(4), planDeadline,
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                } else {
                    setTableCellText(planTable.getRow(rowIndex).getCell(0), "",
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                    setTableCellText(planTable.getRow(rowIndex).getCell(1), "",
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                    setTableCellText(planTable.getRow(rowIndex).getCell(4), "",
                            PLAN_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                }
            }

            if (dataRowsCount > 1) {
                mergeCellsVertically(planTable, 0, 1, dataRowsCount);
                mergeCellsVertically(planTable, 1, 1, dataRowsCount);
                mergeCellsVertically(planTable, 4, 1, dataRowsCount);
            }

            XWPFParagraph spacerAfterPlan = document.createParagraph();
            setParagraphSpacing(spacerAfterPlan);

            addParagraphWithLineBreaks(document,
                    "Представитель заказчика _______________________________________________\n" +
                            "                                                 (Должность, ФИО, контактные данные) ");

            if (!microclimateRows.isEmpty() || !ventilationRows.isEmpty() || hasMedSheet) {
                XWPFParagraph appendixBreak = document.createParagraph();
                setParagraphSpacing(appendixBreak);
                appendixBreak.createRun().addBreak(BreakType.PAGE);

                XWPFParagraph customerAppendixHeader = document.createParagraph();
                customerAppendixHeader.setAlignment(ParagraphAlignment.RIGHT);
                setParagraphSpacing(customerAppendixHeader);
                XWPFRun customerAppendixHeaderRun = customerAppendixHeader.createRun();
                customerAppendixHeaderRun.setFontFamily(FONT_NAME);
                customerAppendixHeaderRun.setFontSize(FONT_SIZE);
                setRunTextWithBreaks(customerAppendixHeaderRun,
                        "Приложения к заявке № " + applicationNumber + "\n" +
                                "Приложение – Данные, предоставлены Заказчиком");

                XWPFParagraph customerAppendixTitle = document.createParagraph();
                customerAppendixTitle.setAlignment(ParagraphAlignment.CENTER);
                setParagraphSpacing(customerAppendixTitle);
                XWPFRun customerAppendixTitleRun = customerAppendixTitle.createRun();
                customerAppendixTitleRun.setText("Данные, предоставленные Заказчиком, за которые он несет ответственность");
                customerAppendixTitleRun.setFontFamily(FONT_NAME);
                customerAppendixTitleRun.setFontSize(FONT_SIZE);
                customerAppendixTitleRun.setBold(true);

                XWPFParagraph spacerBeforeMicroclimate = document.createParagraph();
                setParagraphSpacing(spacerBeforeMicroclimate);

                int sectionIndex = 1;

                if (!microclimateRows.isEmpty()) {
                    XWPFParagraph microclimateTitle = document.createParagraph();
                    setParagraphSpacing(microclimateTitle);
                    XWPFRun microclimateTitleRun = microclimateTitle.createRun();
                    microclimateTitleRun.setFontFamily(FONT_NAME);
                    microclimateTitleRun.setFontSize(FONT_SIZE);
                    microclimateTitleRun.setText(sectionIndex + ".\tДопустимые уровни параметров микроклимата " +
                            "в соответствии с СанПиН 1.2.3685-21 \"Гигиенические нормативы и требования к " +
                            "обеспечению безопасности и (или) безвредности для человека факторов среды обитания\" " +
                            "с указанием места проведения измерений:");

                    int microclimateRowsCount = microclimateRows.size();
                    XWPFTable microclimateTable = document.createTable(1 + microclimateRowsCount, 5);
                    configureTableLayout(microclimateTable, new int[]{6280, 1570, 1570, 1570, 1570});
                    setTableCellText(microclimateTable.getRow(0).getCell(0),
                            "Рабочее место, место проведения измерений, цех, участок,\n" +
                                    "наименование профессии или \nдолжности",
                            MICROCLIMATE_HEADER_MAIN_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(microclimateTable.getRow(0).getCell(1),
                            "Допустимый уровень температуры воздуха, ºС",
                            MICROCLIMATE_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(microclimateTable.getRow(0).getCell(2),
                            "Допустимый уровень результирующей температуры, ºС",
                            MICROCLIMATE_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(microclimateTable.getRow(0).getCell(3),
                            "Допустимый уровень влажности воздуха, %",
                            MICROCLIMATE_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(microclimateTable.getRow(0).getCell(4),
                            "Допустимый уровень скорости движения воздуха, м/с",
                            MICROCLIMATE_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

                    for (int index = 0; index < microclimateRowsCount; index++) {
                        int rowIndex = index + 1;
                        MicroclimateRow row = microclimateRows.get(index);
                        if (row.isSection) {
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(0), row.workplace,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            mergeCellsHorizontally(microclimateTable, rowIndex, 0, 4);
                        } else {
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(0), row.workplace,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(1), row.airTemperature,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(2), row.resultTemperature,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(3), row.humidity,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(microclimateTable.getRow(rowIndex).getCell(4), row.airSpeed,
                                    MICROCLIMATE_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                        }
                    }
                    sectionIndex++;
                }

                if (!ventilationRows.isEmpty()) {
                    XWPFParagraph spacerBeforeVentilation = document.createParagraph();
                    setParagraphSpacing(spacerBeforeVentilation);

                    XWPFParagraph ventilationTitle = document.createParagraph();
                    setParagraphSpacing(ventilationTitle);
                    XWPFRun ventilationTitleRun = ventilationTitle.createRun();
                    ventilationTitleRun.setFontFamily(FONT_NAME);
                    ventilationTitleRun.setFontSize(FONT_SIZE);
                    String ventilationMethodText = ventilationMethod == null ? "" : ventilationMethod.trim();
                    String ventilationMethodClause = ventilationMethodText.isBlank()
                            ? ""
                            : "в соответствии с " + ventilationMethodText + " ";
                    String ventilationTitleText = sectionIndex + ".\tДопустимые уровни скорости движения воздуха, " +
                            "скорости воздушного потока, производительности вентсистем, кратности воздухообмена " +
                            "по притоку и вытяжке " + ventilationMethodClause +
                            "с указанием места проведения измерений:";
                    ventilationTitleRun.setText(ventilationTitleText.trim());

                    int ventilationRowsCount = ventilationRows.size();
                    XWPFTable ventilationTable = document.createTable(1 + ventilationRowsCount, 4);
                    configureTableLayout(ventilationTable, new int[]{6280, 1570, 1570, 1570});
                    setTableCellText(ventilationTable.getRow(0).getCell(0),
                            "Рабочее место, место проведения измерений (Приток/вытяжка)",
                            VENTILATION_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(ventilationTable.getRow(0).getCell(1),
                            "Объем помещения, м^3",
                            VENTILATION_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(ventilationTable.getRow(0).getCell(2),
                            "Допустимый уровень производительности венсистем, м^3/ч",
                            VENTILATION_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);
                    setTableCellText(ventilationTable.getRow(0).getCell(3),
                            "Допустимый уровень кратности воздухообмена, ч^-1",
                            VENTILATION_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

                    for (int index = 0; index < ventilationRowsCount; index++) {
                        int rowIndex = index + 1;
                        VentilationRow row = ventilationRows.get(index);
                        if (row.isSection) {
                            setTableCellText(ventilationTable.getRow(rowIndex).getCell(0), row.workplace,
                                    VENTILATION_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            mergeCellsHorizontally(ventilationTable, rowIndex, 0, 3);
                        } else {
                            setTableCellText(ventilationTable.getRow(rowIndex).getCell(0), row.workplace,
                                    VENTILATION_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(ventilationTable.getRow(rowIndex).getCell(1), row.volume,
                                    VENTILATION_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(ventilationTable.getRow(rowIndex).getCell(2), row.performance,
                                    VENTILATION_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                            setTableCellText(ventilationTable.getRow(rowIndex).getCell(3), row.airExchange,
                                    VENTILATION_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                        }
                    }
                    sectionIndex++;
                }

                if (hasMedSheet) {
                    XWPFParagraph spacerBeforeMed = document.createParagraph();
                    setParagraphSpacing(spacerBeforeMed);

                    XWPFParagraph medTitle = document.createParagraph();
                    setParagraphSpacing(medTitle);
                    XWPFRun medTitleRun = medTitle.createRun();
                    medTitleRun.setFontFamily(FONT_NAME);
                    medTitleRun.setFontSize(FONT_SIZE);
                    String medMethodText = medNormativeMethod == null ? "" : medNormativeMethod.trim();
                    String medMethodClause = medMethodText.isBlank()
                            ? ""
                            : "в соответствии с " + medMethodText + " ";
                    String medTitleText = sectionIndex + ".\tДопустимый уровень мощности дозы гамма-излучения "
                            + medMethodClause
                            + "Превышение мощности дозы, измеренной на открытой местности, не более чем на 0,3 мкЗв/ч";
                    medTitleRun.setText(medTitleText.trim());

                    if (medRows.isEmpty()) {
                        medRows.add("");
                    }
                    int medRowsCount = medRows.size();
                    XWPFTable medTable = document.createTable(1 + medRowsCount, 1);
                    configureTableLayout(medTable, new int[]{12560});
                    setTableCellText(medTable.getRow(0).getCell(0),
                            "Наименование места проведения измерений",
                            MED_TABLE_FONT_SIZE, true, ParagraphAlignment.CENTER);

                    for (int index = 0; index < medRowsCount; index++) {
                        int rowIndex = index + 1;
                        String row = medRows.get(index);
                        setTableCellText(medTable.getRow(rowIndex).getCell(0), row,
                                MED_TABLE_FONT_SIZE, false, ParagraphAlignment.LEFT);
                    }
                    sectionIndex++;
                }
            }

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveRequestFormFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), REQUEST_FORM_NAME);
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
        setHeaderCellText(table.getRow(0).getCell(1), "Форма заявки Ф46 ДП ИЛ 2-2023");
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

    private static void configureTableLayout(XWPFTable table, int[] columnWidths) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(scaleWidth(12560)));

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

    private static void setTableCellText(XWPFTableCell cell, String text, boolean bold,
                                         ParagraphAlignment alignment) {
        setTableCellText(cell, text, FONT_SIZE, bold, alignment);
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize, boolean bold,
                                         ParagraphAlignment alignment) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(alignment);
        setParagraphSpacing(paragraph);
        XWPFRun run = paragraph.createRun();
        setRunTextWithBreaks(run, text);
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);
        run.setBold(bold);
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

    private static String resolveApplicationNumberFromMap(File mapFile) {
        String line = readMapRowText(mapFile, MAP_APPLICATION_ROW_INDEX);
        String lowerLine = line.toLowerCase(Locale.ROOT);
        int applicationIndex = lowerLine.indexOf("заявка");
        if (applicationIndex >= 0) {
            String value = line.substring(applicationIndex + "заявка".length()).trim();
            return trimLeadingPunctuation(value);
        }
        int commaIndex = line.indexOf(',');
        if (commaIndex >= 0 && commaIndex + 1 < line.length()) {
            return line.substring(commaIndex + 1).trim();
        }
        return line.trim();
    }

    private static CustomerInfo resolveCustomerInfo(File mapFile) {
        String line = readMapRowText(mapFile, MAP_CUSTOMER_ROW_INDEX);
        String value = extractValueAfterPrefix(line, "1. Заказчик");
        if (value.isBlank()) {
            value = extractValueAfterPrefix(line, "Заказчик");
        }
        List<String> parts = splitCommaParts(value);
        String name = parts.isEmpty() ? "" : parts.get(0);
        String email = parts.size() > 1 ? parts.get(1) : "";
        String phone = "";
        if (parts.size() > 3) {
            phone = parts.get(3);
        } else if (parts.size() > 2) {
            phone = parts.get(2);
        }
        return new CustomerInfo(name, email, phone);
    }

    private static List<String> splitCommaParts(String value) {
        List<String> parts = new ArrayList<>();
        if (value == null || value.isBlank()) {
            return parts;
        }
        for (String part : value.split(",")) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) {
                parts.add(trimmed);
            }
        }
        return parts;
    }

    private static String readMapRowText(File mapFile, int rowIndex) {
        if (mapFile == null || !mapFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            String firstValue = "";
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.isEmpty()) {
                    continue;
                }
                if (firstValue.isEmpty()) {
                    firstValue = text;
                }
                if (text.contains("Номер протокола")
                        || text.contains("Заказчик")
                        || text.toLowerCase(Locale.ROOT).contains("договор")
                        || text.toLowerCase(Locale.ROOT).contains("заявка")) {
                    return text;
                }
            }
            return firstValue;
        } catch (Exception ex) {
            return "";
        }
    }

    private static String extractValueAfterPrefix(String line, String prefix) {
        if (line == null) {
            return "";
        }
        String normalized = line.trim();
        int index = normalized.indexOf(prefix);
        if (index >= 0) {
            return normalized.substring(index + prefix.length()).replace(":", "").trim();
        }
        return normalized.replace(":", "").trim();
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

    private static String resolveObjectName(File mapFile) {
        String value = findValueByPrefix(mapFile, OBJECT_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "4. Наименование объекта");
        }
        return value;
    }

    private static String resolveObjectAddress(File mapFile) {
        String value = findValueByPrefix(mapFile, ADDRESS_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "Адрес объекта:");
        }
        return value;
    }

    private static String findValueByPrefix(File mapFile, String prefix) {
        if (mapFile == null || !mapFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = formatter.formatCellValue(cell).trim();
                    if (text.startsWith(prefix)) {
                        String tail = text.substring(prefix.length()).trim();
                        if (!tail.isEmpty()) {
                            return trimLeadingPunctuation(tail);
                        }
                        Cell next = row.getCell(cell.getColumnIndex() + 1);
                        if (next != null) {
                            String nextText = formatter.formatCellValue(next).trim();
                            if (!nextText.isEmpty()) {
                                return trimLeadingPunctuation(nextText);
                            }
                        }
                    }
                }
            }
        } catch (Exception ignored) {
            return "";
        }
        return "";
    }

    private static String resolveProtocolDateFromSource(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            Row row = sheet.getRow(6);
            if (row == null) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            String lastValue = "";
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.isEmpty()) {
                    if (!lastValue.isEmpty()) {
                        break;
                    }
                    continue;
                }
                lastValue = text;
            }
            return lastValue;
        } catch (Exception ignored) {
            return "";
        }
    }

    private static List<MicroclimateRow> resolveMicroclimateRows(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new ArrayList<>();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = findSheetByName(workbook, "Микроклимат");
            if (sheet == null) {
                return new ArrayList<>();
            }
            List<MicroclimateRow> rows = new ArrayList<>();
            DataFormatter formatter = new DataFormatter();
            int lastRow = sheet.getLastRowNum();
            for (int rowIndex = 5; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                String sectionText = formatter.formatCellValue(row.getCell(0)).trim();
                String workplace = formatter.formatCellValue(row.getCell(1)).trim();
                String airTemperature = formatter.formatCellValue(row.getCell(8)).trim();
                String resultTemperature = formatter.formatCellValue(row.getCell(13)).trim();
                String humidity = formatter.formatCellValue(row.getCell(17)).trim();
                String airSpeed = formatter.formatCellValue(row.getCell(21)).trim();

                boolean isSection = !sectionText.isEmpty()
                        && workplace.isEmpty()
                        && airTemperature.isEmpty()
                        && resultTemperature.isEmpty()
                        && humidity.isEmpty()
                        && airSpeed.isEmpty();
                if (isSection) {
                    rows.add(MicroclimateRow.section(sectionText));
                    continue;
                }

                if (workplace.isEmpty()) {
                    continue;
                }

                int scanIndex = rowIndex + 1;
                while (scanIndex <= lastRow
                        && (airTemperature.isEmpty() || resultTemperature.isEmpty()
                        || humidity.isEmpty() || airSpeed.isEmpty())) {
                    Row scanRow = sheet.getRow(scanIndex);
                    if (scanRow == null) {
                        scanIndex++;
                        continue;
                    }
                    String nextSection = formatter.formatCellValue(scanRow.getCell(0)).trim();
                    String nextWorkplace = formatter.formatCellValue(scanRow.getCell(1)).trim();
                    boolean nextIsSection = !nextSection.isEmpty()
                            && nextWorkplace.isEmpty()
                            && formatter.formatCellValue(scanRow.getCell(8)).trim().isEmpty()
                            && formatter.formatCellValue(scanRow.getCell(13)).trim().isEmpty()
                            && formatter.formatCellValue(scanRow.getCell(17)).trim().isEmpty()
                            && formatter.formatCellValue(scanRow.getCell(21)).trim().isEmpty();
                    if (!nextWorkplace.isEmpty() || nextIsSection) {
                        break;
                    }
                    if (airTemperature.isEmpty()) {
                        airTemperature = formatter.formatCellValue(scanRow.getCell(8)).trim();
                    }
                    if (resultTemperature.isEmpty()) {
                        resultTemperature = formatter.formatCellValue(scanRow.getCell(13)).trim();
                    }
                    if (humidity.isEmpty()) {
                        humidity = formatter.formatCellValue(scanRow.getCell(17)).trim();
                    }
                    if (airSpeed.isEmpty()) {
                        airSpeed = formatter.formatCellValue(scanRow.getCell(21)).trim();
                    }
                    scanIndex++;
                }

                rows.add(new MicroclimateRow(workplace, airTemperature, resultTemperature, humidity, airSpeed));
            }
            if (rows.isEmpty()) {
                rows.add(new MicroclimateRow("", "", "", "", ""));
            }
            return rows;
        } catch (Exception ignored) {
            return new ArrayList<>();
        }
    }

    private static List<VentilationRow> resolveVentilationRows(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new ArrayList<>();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = findSheetByKeyword(workbook, "вентиляция");
            if (sheet == null) {
                return new ArrayList<>();
            }
            List<VentilationRow> rows = new ArrayList<>();
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            int lastRow = sheet.getLastRowNum();
            int emptyStreak = 0;
            boolean started = false;

            for (int rowIndex = VENTILATION_SOURCE_START_ROW; rowIndex <= lastRow; rowIndex++) {
                CellRangeAddress mergedRow = findMergedRegion(sheet, rowIndex, 0);
                if (isVentilationMergedRow(mergedRow, rowIndex)) {
                    String sectionText = readMergedCellValue(sheet, rowIndex, 0, formatter, evaluator).trim();
                    if (!sectionText.isEmpty()) {
                        rows.add(VentilationRow.section(sectionText));
                        started = true;
                    }
                    emptyStreak = 0;
                    continue;
                }

                String workplace = readMergedCellValue(sheet, rowIndex, 2, formatter, evaluator).trim();
                String volume = readMergedCellValue(sheet, rowIndex, 8, formatter, evaluator).trim();
                String performance = readMergedCellValue(sheet, rowIndex, 10, formatter, evaluator).trim();
                String airExchange = readMergedCellValue(sheet, rowIndex, 11, formatter, evaluator).trim();

                if (workplace.isEmpty() && volume.isEmpty() && performance.isEmpty() && airExchange.isEmpty()) {
                    emptyStreak++;
                    if (started && emptyStreak >= 20) {
                        break;
                    }
                    continue;
                }
                emptyStreak = 0;
                started = true;
                rows.add(new VentilationRow(workplace, volume, performance, airExchange));
            }

            if (rows.isEmpty()) {
                rows.add(new VentilationRow("", "", "", ""));
            }
            return rows;
        } catch (Exception ignored) {
            return new ArrayList<>();
        }
    }

    private static String resolveVentilationNormativeMethod(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int headerRowIndex = findNormativeHeaderRow(sheet, formatter);
            if (headerRowIndex < 0) {
                return "";
            }
            for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }
                String indicator = findFirstValueInRange(row, formatter, 0, 4);
                if (indicator.isEmpty()) {
                    continue;
                }
                String normalized = indicator.toLowerCase(Locale.ROOT);
                if (normalized.contains("скорость воздушного потока")) {
                    return resolveNormativeMethodValue(row, formatter);
                }
            }
            return "";
        } catch (Exception ignored) {
            return "";
        }
    }

    private static List<String> resolveMedRows(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new ArrayList<>();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            Sheet sheet = findSheetByName(workbook, "МЭД (2)");
            if (sheet == null) {
                return new ArrayList<>();
            }
            List<String> rows = new ArrayList<>();
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            int lastRow = sheet.getLastRowNum();
            int emptyStreak = 0;
            boolean started = false;

            for (int rowIndex = MED_SOURCE_START_ROW; rowIndex <= lastRow; rowIndex++) {
                CellRangeAddress mergedRow = findMergedRegion(sheet, rowIndex, 0);
                if (isMedMergedRow(mergedRow, rowIndex)) {
                    String text = readMergedCellValue(sheet, rowIndex, 0, formatter, evaluator).trim();
                    if (text.isEmpty() && !hasRowContentInRange(sheet, rowIndex, formatter, evaluator, 0, 5)) {
                        if (started) {
                            break;
                        }
                        continue;
                    }
                    rows.add(text);
                    started = true;
                    emptyStreak = 0;
                    continue;
                }

                String place = readMergedCellValue(sheet, rowIndex, 1, formatter, evaluator).trim();
                if (place.isEmpty()) {
                    emptyStreak++;
                    if (started && emptyStreak >= 10) {
                        break;
                    }
                    continue;
                }
                emptyStreak = 0;
                started = true;
                rows.add(place);
            }

            return rows;
        } catch (Exception ignored) {
            return new ArrayList<>();
        }
    }

    private static String resolveMedNormativeMethod(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int headerRowIndex = findNormativeHeaderRow(sheet, formatter);
            if (headerRowIndex < 0) {
                return "";
            }
            for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }
                String indicator = findFirstValueInRange(row, formatter, 0, 4);
                if (indicator.isEmpty()) {
                    continue;
                }
                String normalized = indicator.toLowerCase(Locale.ROOT);
                if (normalized.contains("мощность дозы гамма-излучения")) {
                    return resolveNormativeMethodValue(row, formatter);
                }
            }
            return "";
        } catch (Exception ignored) {
            return "";
        }
    }

    private static Sheet findSheetByName(Workbook workbook, String sheetName) {
        if (workbook == null) {
            return null;
        }
        int sheetCount = workbook.getNumberOfSheets();
        for (int index = 0; index < sheetCount; index++) {
            Sheet sheet = workbook.getSheetAt(index);
            if (sheet != null && sheet.getSheetName().equalsIgnoreCase(sheetName)) {
                return sheet;
            }
        }
        return null;
    }

    private static boolean hasSheetByName(File sourceFile, String sheetName) {
        if (sourceFile == null || !sourceFile.exists()) {
            return false;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            return findSheetByName(workbook, sheetName) != null;
        } catch (Exception ignored) {
            return false;
        }
    }

    private static Sheet findSheetByKeyword(Workbook workbook, String keyword) {
        if (workbook == null || keyword == null) {
            return null;
        }
        int sheetCount = workbook.getNumberOfSheets();
        for (int index = 0; index < sheetCount; index++) {
            Sheet sheet = workbook.getSheetAt(index);
            if (sheet == null) {
                continue;
            }
            String name = sheet.getSheetName();
            if (name != null && name.toLowerCase(Locale.ROOT).contains(keyword.toLowerCase(Locale.ROOT))) {
                return sheet;
            }
        }
        return null;
    }

    private static List<NormativeRow> resolveNormativeRows(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new ArrayList<>();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return new ArrayList<>();
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int headerRowIndex = findNormativeHeaderRow(sheet, formatter);
            if (headerRowIndex < 0) {
                return new ArrayList<>();
            }
            List<NormativeRow> rows = new ArrayList<>();
            for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }
                String indicator = findFirstValueInRange(row, formatter, 0, 4);
                String method = findFirstValueInRange(row, formatter, 15, 25);
                if (indicator.contains(NORMATIVE_HEADER_TITLE)) {
                    continue;
                }
                if (indicator.isEmpty() && method.isEmpty()) {
                    break;
                }
                rows.add(new NormativeRow(indicator, method));
            }
            return rows;
        } catch (Exception ignored) {
            return new ArrayList<>();
        }
    }

    private static int findNormativeHeaderRow(Sheet sheet, DataFormatter formatter) {
        int sectionStart = -1;
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (sectionStart < 0 && text.contains(NORMATIVE_SECTION_TITLE)) {
                    sectionStart = row.getRowNum();
                }
                if (sectionStart >= 0 && text.contains(NORMATIVE_HEADER_TITLE)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static String findFirstValueInRange(Row row, DataFormatter formatter, int startCol, int endCol) {
        for (int col = startCol; col <= endCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty()) {
                return text;
            }
        }
        return "";
    }

    private static String resolveNormativeMethodValue(Row row, DataFormatter formatter) {
        String method = findFirstValueInRange(row, formatter, 5, 14);
        if (!method.isEmpty()) {
            return method;
        }
        return findFirstValueInRange(row, formatter, 15, 25);
    }

    private static boolean isVentilationMergedRow(CellRangeAddress region, int rowIndex) {
        return region != null
                && region.getFirstRow() == rowIndex
                && region.getLastRow() == rowIndex
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= 11;
    }

    private static boolean isMedMergedRow(CellRangeAddress region, int rowIndex) {
        return region != null
                && region.getFirstRow() == rowIndex
                && region.getLastRow() == rowIndex
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= 5;
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) {
            return null;
        }
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region != null && region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String readMergedCellValue(Sheet sheet,
                                              int rowIndex,
                                              int colIndex,
                                              DataFormatter formatter,
                                              FormulaEvaluator evaluator) {
        if (sheet == null) {
            return "";
        }
        CellRangeAddress merged = findMergedRegion(sheet, rowIndex, colIndex);
        int targetRow = rowIndex;
        int targetCol = colIndex;
        if (merged != null) {
            targetRow = merged.getFirstRow();
            targetCol = merged.getFirstColumn();
        }
        Row row = sheet.getRow(targetRow);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(targetCol);
        if (cell == null) {
            return "";
        }
        String value = formatter.formatCellValue(cell, evaluator);
        return value == null ? "" : value.trim();
    }

    private static boolean hasRowContentInRange(Sheet sheet,
                                                int rowIndex,
                                                DataFormatter formatter,
                                                FormulaEvaluator evaluator,
                                                int startCol,
                                                int endCol) {
        if (sheet == null) {
            return false;
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return false;
        }
        for (int col = startCol; col <= endCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            String text = formatter.formatCellValue(cell, evaluator);
            if (text != null && !text.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private record CustomerInfo(String name, String email, String phone) {
    }

    private static final class NormativeRow {
        private final String indicator;
        private final String method;

        private NormativeRow(String indicator, String method) {
            this.indicator = indicator == null ? "" : indicator;
            this.method = method == null ? "" : method;
        }
    }

    private static final class MicroclimateRow {
        private final String workplace;
        private final String airTemperature;
        private final String resultTemperature;
        private final String humidity;
        private final String airSpeed;
        private final boolean isSection;

        private MicroclimateRow(String workplace, String airTemperature, String resultTemperature, String humidity,
                                String airSpeed) {
            this.workplace = workplace == null ? "" : workplace;
            this.airTemperature = airTemperature == null ? "" : airTemperature;
            this.resultTemperature = resultTemperature == null ? "" : resultTemperature;
            this.humidity = humidity == null ? "" : humidity;
            this.airSpeed = airSpeed == null ? "" : airSpeed;
            this.isSection = false;
        }

        private MicroclimateRow(String section) {
            this.workplace = section == null ? "" : section;
            this.airTemperature = "";
            this.resultTemperature = "";
            this.humidity = "";
            this.airSpeed = "";
            this.isSection = true;
        }

        private static MicroclimateRow section(String section) {
            return new MicroclimateRow(section);
        }
    }

    private static final class VentilationRow {
        private final String workplace;
        private final String volume;
        private final String performance;
        private final String airExchange;
        private final boolean isSection;

        private VentilationRow(String workplace, String volume, String performance, String airExchange) {
            this.workplace = workplace == null ? "" : workplace;
            this.volume = volume == null ? "" : volume;
            this.performance = performance == null ? "" : performance;
            this.airExchange = airExchange == null ? "" : airExchange;
            this.isSection = false;
        }

        private VentilationRow(String section) {
            this.workplace = section == null ? "" : section;
            this.volume = "";
            this.performance = "";
            this.airExchange = "";
            this.isSection = true;
        }

        private static VentilationRow section(String section) {
            return new VentilationRow(section);
        }
    }
}
