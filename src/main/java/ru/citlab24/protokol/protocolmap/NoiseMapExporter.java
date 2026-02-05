package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

public final class NoiseMapExporter {
    private static final String REGISTRATION_PREFIX = "Регистрационный номер карты замеров:";
    private static final double COLUMN_WIDTH_SCALE = 0.9;
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final int TITLE_MEASUREMENT_DATES_ROW = 7;
    private static final String PRIMARY_FOLDER_NAME = "Первичка Шумы";

    private NoiseMapExporter() {
    }

    public static File generateMap(File sourceFile, String workDeadline, String customerInn) throws IOException {
        String registrationNumber = resolveRegistrationNumber(sourceFile);
        MapHeaderData headerData = resolveHeaderData(sourceFile);
        String protocolNumber = resolveProtocolNumber(sourceFile);
        String contractText = resolveContractText(sourceFile);
        String measurementPerformer = resolveMeasurementPerformer(sourceFile);
        String titleMeasurementDates = resolveTitleMeasurementDates(sourceFile);
        String measurementDates = titleMeasurementDates.isBlank()
                ? resolveMeasurementDates(sourceFile)
                : titleMeasurementDates;
        String controlDate = resolveControlDate(sourceFile);
        String additionalInfo = resolveAdditionalInfo(sourceFile);
        String specialConditions = resolveSpecialConditions(sourceFile);
        String measurementMethods = resolveMeasurementMethods(sourceFile);
        java.util.List<InstrumentData> instruments = resolveMeasurementInstruments(sourceFile);
        File targetFile = buildTargetFile(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("карта замеров");
            applySheetDefaults(workbook, sheet);
            applyHeaders(sheet, registrationNumber);
            createTitleRows(workbook, sheet, registrationNumber, headerData, measurementPerformer, measurementDates, controlDate);
            createSecondPageRows(workbook, sheet, protocolNumber, contractText, headerData,
                    additionalInfo, specialConditions, measurementMethods, instruments);
            java.util.List<String> measurementDatesList = extractMeasurementDatesList(measurementDates);
            createProtocolTabs(sourceFile, workbook, registrationNumber, measurementDatesList);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                workbook.write(out);
            }
        }

        ProtocolIssuanceSheetExporter.generate(sourceFile, targetFile);
        MeasurementCardRegistrationSheetExporter.generate(sourceFile, targetFile);

        MeasurementPlanExporter.generate(sourceFile, targetFile, workDeadline);
        RequestFormExporter.generateForNoise(sourceFile, targetFile, workDeadline, customerInn);
        RequestAnalysisSheetExporter.generate(targetFile);

        return targetFile;

    }

    private static String resolveRegistrationNumber(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findRegistrationNumber(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findRegistrationNumber(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.startsWith(REGISTRATION_PREFIX)) {
                    String tail = text.substring(REGISTRATION_PREFIX.length()).trim();
                    if (!tail.isEmpty()) {
                        return tail;
                    }
                    Cell next = row.getCell(cell.getColumnIndex() + 1);
                    if (next != null) {
                        String nextText = formatter.formatCellValue(next).trim();
                        if (!nextText.isEmpty()) {
                            return nextText;
                        }
                    }
                }
            }
        }
        return "";
    }

    private static File buildTargetFile(File sourceFile) {
        String name = sourceFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = dotIndex > 0 ? name.substring(0, dotIndex) : name;
        File primaryFolder = ensurePrimaryFolder(sourceFile);
        return new File(primaryFolder, baseName + "_карта.xlsx");
    }

    private static File ensurePrimaryFolder(File sourceFile) {
        File parent = sourceFile != null ? sourceFile.getParentFile() : null;
        if (parent == null) {
            parent = new File(".");
        }
        File primaryFolder = new File(parent, PRIMARY_FOLDER_NAME);
        if (!primaryFolder.exists()) {
            primaryFolder.mkdirs();
        }
        return primaryFolder;
    }

    private static void applySheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 12);
        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);

        int[] widthsPx = buildColumnWidthsPx();
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
            sheet.setDefaultColumnStyle(col, baseStyle);
        }

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));
    }

    private static void applyHeaders(Sheet sheet, String registrationNumber) {
        String font = "&\"Arial\"&12";
        Header header = sheet.getHeader();
        header.setLeft(font + "Испытательная лаборатория\nООО «ЦИТ»");
        header.setCenter(font + "Карта замеров № " + registrationNumber + "\nФ8 РИ ИЛ 2-2023");
        header.setRight(font + "\nКоличество страниц: &[Страница] / &[Страниц] \n ");
    }

    private static void createTitleRows(Workbook workbook,
                                        Sheet sheet,
                                        String registrationNumber,
                                        MapHeaderData headerData,
                                        String measurementPerformer,
                                        String measurementDates,
                                        String controlDate) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        Font sectionFont = workbook.createFont();
        sectionFont.setFontName("Arial");
        sectionFont.setFontHeightInPoints((short) 14);
        sectionFont.setBold(true);

        CellStyle sectionStyle = workbook.createCellStyle();
        sectionStyle.setFont(sectionFont);
        sectionStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        sectionStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        sectionStyle.setWrapText(true);

        Font sectionValueFont = workbook.createFont();
        sectionValueFont.setFontName("Arial");
        sectionValueFont.setFontHeightInPoints((short) 12);
        sectionValueFont.setBold(false);

        CellStyle sectionMixedStyle = workbook.createCellStyle();
        sectionMixedStyle.setFont(sectionValueFont);
        sectionMixedStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        sectionMixedStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        sectionMixedStyle.setWrapText(true);

        setMergedCellValue(sheet, 0, "Испытательная лаборатория Общества с ограниченной ответственностью", titleStyle);

        setMergedCellValue(sheet, 1, "«Центр исследовательских технологий»", titleStyle);

        Row spacerRow = sheet.createRow(2);
        spacerRow.setHeightInPoints(pixelsToPoints(3));

        setMergedCellValue(sheet, 3, "КАРТА ЗАМЕРОВ № " + registrationNumber, titleStyle);

        sheet.createRow(4);

        String customerPrefix = "1. Заказчик: ";
        String customerValue = safe(headerData.customerNameAndContacts);
        setMergedCellValueWithPrefix(sheet, 5, customerPrefix, customerValue, sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 5, 0, 31, customerPrefix + customerValue);

        Row heightRow = sheet.createRow(6);
        heightRow.setHeightInPoints(pixelsToPoints(16));

        String datesPrefix = "2. Дата замеров: ";
        String datesValue = safe(measurementDates);
        setMergedCellValueWithPrefix(sheet, 7, datesPrefix, datesValue, sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 7, 0, 31, datesPrefix + datesValue);

        Row row8 = sheet.createRow(8);
        row8.setHeightInPoints(pixelsToPoints(16));

        String performerPrefix = "3. Измерения провел, подпись: ";
        String performerValue = safe(measurementPerformer);
        setMergedCellValueWithPrefix(sheet, 9, performerPrefix, performerValue,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row row10 = sheet.createRow(10);
        row10.setHeightInPoints(pixelsToPoints(16));

        String representativePrefix = "4. Измерения проведены в присутствии представителя: ";
        String representativeValue = safe(headerData.representative);
        setMergedCellValueWithPrefix(sheet, 11, representativePrefix, representativeValue,
                sectionFont, sectionValueFont, sectionMixedStyle);
        adjustRowHeightForMergedTextDoubling(sheet, 11, 0, 31, representativePrefix + representativeValue);

        Row row12 = sheet.createRow(12);
        row12.setHeightInPoints(pixelsToPoints(16));

        setMergedCellValue(sheet, 13, "Подпись:_______________________________", sectionStyle);

        Row row14 = sheet.createRow(14);
        row14.setHeightInPoints(pixelsToPoints(16));

        String controlPrefix = "5. Контроль ведения записей осуществлен: ";
        String controlSuffix = resolveControlSuffix(measurementPerformer);
        setMergedCellValueWithPrefix(sheet, 15, controlPrefix, controlSuffix,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row row16 = sheet.createRow(16);
        row16.setHeightInPoints(pixelsToPoints(16));

        String controlResultPrefix = "Результат контроля: ";
        String controlResultValue = "соответствует/не соответствует";
        setMergedCellValueWithPrefix(sheet, 17, controlResultPrefix, controlResultValue,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row row18 = sheet.createRow(18);
        row18.setHeightInPoints(pixelsToPoints(16));

        String controlDatePrefix = "Дата контроля: ";
        setMergedCellValueWithPrefix(sheet, 19, controlDatePrefix, controlDate,
                sectionFont, sectionValueFont, sectionMixedStyle);

        Row spacerAfterControlDate = sheet.createRow(20);
        spacerAfterControlDate.setHeightInPoints(pixelsToPoints(16));
    }

    private static void createSecondPageRows(Workbook workbook,
                                             Sheet sheet,
                                             String protocolNumber,
                                             String contractText,
                                             MapHeaderData headerData,
                                             String additionalInfo,
                                             String specialConditions,
                                             String measurementMethods,
                                             java.util.List<InstrumentData> instruments) {
        int startRow = 21;
        sheet.setRowBreak(startRow - 1);

        Font plainFont = workbook.createFont();
        plainFont.setFontName("Arial");
        plainFont.setFontHeightInPoints((short) 12);
        plainFont.setBold(false);

        CellStyle plainStyle = workbook.createCellStyle();
        plainStyle.setFont(plainFont);
        plainStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        plainStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        plainStyle.setWrapText(true);

        int rowIndex = startRow;
        setMergedCellValue(sheet, rowIndex,
                "1. Номер протокола " + safe(protocolNumber), plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "2. Договор " + safe(contractText), plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "3. Наименование и контактные данные Заказчика: " + safe(headerData.customerNameAndContacts),
                plainStyle);
        rowIndex++;
        setMergedCellValue(sheet, rowIndex,
                "Юридический адрес заказчика: " + safe(headerData.customerLegalAddress), plainStyle);
        rowIndex++;
        String objectNameText = "4. Наименование объекта: " + safe(headerData.objectName);
        setMergedCellValue(sheet, rowIndex, objectNameText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, objectNameText);
        rowIndex++;

        String objectAddressText = "Адрес объекта " + safe(headerData.objectAddress);
        setMergedCellValue(sheet, rowIndex, objectAddressText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, objectAddressText);
        rowIndex++;

        String additionalInfoText = "5. Дополнительные сведения (" + safe(additionalInfo) + ")";
        setMergedCellValue(sheet, rowIndex, additionalInfoText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, additionalInfoText);
        rowIndex++;

        String specialConditionsText = "5.1. Особые условия: " + safe(specialConditions);
        setMergedCellValue(sheet, rowIndex, specialConditionsText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, specialConditionsText);
        rowIndex++;

        String measurementMethodsText = "5.2. Методы измерения " + safe(measurementMethods);
        setMergedCellValue(sheet, rowIndex, measurementMethodsText, plainStyle);
        adjustRowHeightForMergedText(sheet, rowIndex, 0, 31, measurementMethodsText);
        addRowHeightPixels(sheet, rowIndex, 20);
        rowIndex++;

        setMergedCellValue(sheet, rowIndex,
                "5.3. Приборы для измерения (используемое отметить):", plainStyle);
        rowIndex++;

        Font tableHeaderFont = workbook.createFont();
        tableHeaderFont.setFontName("Arial");
        tableHeaderFont.setFontHeightInPoints((short) 12);
        tableHeaderFont.setBold(true);

        CellStyle tableHeaderStyle = workbook.createCellStyle();
        tableHeaderStyle.setFont(tableHeaderFont);
        tableHeaderStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        tableHeaderStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        tableHeaderStyle.setWrapText(true);
        setThinBorders(tableHeaderStyle);

        CellStyle tableCellStyle = workbook.createCellStyle();
        tableCellStyle.setFont(plainFont);
        tableCellStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        tableCellStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        tableCellStyle.setWrapText(true);
        setThinBorders(tableCellStyle);

        CellStyle checkboxStyle = workbook.createCellStyle();
        checkboxStyle.setFont(plainFont);
        checkboxStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        checkboxStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(checkboxStyle);

        rowIndex = addInstrumentRow(sheet, rowIndex, "Наименование", "зав. №", "☑",
                tableHeaderStyle, tableHeaderStyle, checkboxStyle);

        if (instruments != null) {
            for (InstrumentData instrument : instruments) {
                rowIndex = addInstrumentRow(sheet, rowIndex, instrument.name, instrument.serialNumber, "☑",
                        tableCellStyle, tableCellStyle, checkboxStyle);
            }
        }

        Row spacerAfterInstruments = sheet.createRow(rowIndex);
        spacerAfterInstruments.setHeightInPoints(pixelsToPoints(16));
        rowIndex++;

        sheet.setRowBreak(rowIndex - 1);

        setMergedCellValue(sheet, rowIndex,
                "5.4 калибровочный уровень: 94 дБ\n" +
                        "показания шумомера перед измерениями на частоте 1 кГц:\n" +
                        "показания шумомера после измерений на частоте 1 кГц:",
                plainStyle);
        Row calibrationRow = sheet.getRow(rowIndex);
        if (calibrationRow != null) {
            calibrationRow.setHeightInPoints(sheet.getDefaultRowHeightInPoints() * 3);
        }
        rowIndex++;
    }

    private static void createProtocolTabs(File sourceFile,
                                           Workbook workbook,
                                           String registrationNumber,
                                           java.util.List<String> measurementDates) {
        if (sourceFile == null || !sourceFile.exists()) {
            return;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {
            int sheetCount = sourceWorkbook.getNumberOfSheets();
            java.util.List<String> sheetNames = new java.util.ArrayList<>();
            for (int idx = 1; idx < sheetCount; idx++) {
                Sheet sourceSheet = sourceWorkbook.getSheetAt(idx);
                if (sourceSheet == null) {
                    continue;
                }
                String sheetName = sourceSheet.getSheetName();
                if (isGeneratorSheet(sheetName) || isIgnoredProtocolSheet(sheetName)) {
                    continue;
                }
                if (!isMicroclimateSheet(sheetName)) {
                    sheetNames.add(sheetName);
                }
            }

            if (workbook.getSheet("Микроклимат") == null) {
                PhysicalFactorsMapResultsTabBuilder.createResultsSheet(workbook,
                        measurementDates,
                        false);
                Sheet microclimateSheet = workbook.getSheet("Микроклимат");
                if (microclimateSheet != null) {
                    applyHeaders(microclimateSheet, registrationNumber);
                }
            }

            boolean noiseHeaderApplied = false;
            for (String sheetName : sheetNames) {
                Sheet sourceSheet = sourceWorkbook.getSheet(sheetName);
                Sheet targetSheet = workbook.createSheet(sheetName);
                applyNoiseResultsSheetDefaults(workbook, targetSheet);
                applyHeaders(targetSheet, registrationNumber);
                boolean hasFullHeader = false;
                if (!noiseHeaderApplied && isNoiseProtocolSheet(sheetName)) {
                    buildNoiseResultsHeader(workbook, targetSheet);
                    noiseHeaderApplied = true;
                    hasFullHeader = true;
                } else {
                    addSimpleNumberingRow(workbook, targetSheet);
                }
                fillNoiseResultsFromProtocol(sourceSheet, workbook, targetSheet, hasFullHeader);
            }
        } catch (Exception ex) {
            // Игнорируем ошибки чтения, чтобы карта всё равно сформировалась.
        }
    }

    private static void applyNoiseResultsSheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 9);

        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        baseStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        baseStyle.setWrapText(true);

        int[] widthsPx = buildNoiseResultsColumnWidthsPx();
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
            sheet.setDefaultColumnStyle(col, baseStyle);
        }

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));

        // Excel 2007+ (.xlsx): 1 048 576 строк => последний индекс 1 048 575
        int maxRow = 1_048_575;
        workbook.setPrintArea(workbook.getSheetIndex(sheet), 0, 22, 0, maxRow);
    }


    private static int[] buildNoiseResultsColumnWidthsPx() {
        return new int[]{
                28, 28, 160, 130, 25, 25, 25, 25, 25, 34, 34, 34, 34, 34, 34,
                34, 34, 34, 34, 26, 11, 20, 43
        };
    }

    private static void buildNoiseResultsHeader(Workbook workbook, Sheet sheet) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 9);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
        titleStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFont(titleFont);
        headerStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        headerStyle.setWrapText(true);
        setThinBorders(headerStyle);

        Font smallFont = workbook.createFont();
        smallFont.setFontName("Arial");
        smallFont.setFontHeightInPoints((short) 8);

        CellStyle headerSmallStyle = workbook.createCellStyle();
        headerSmallStyle.cloneStyleFrom(headerStyle);
        headerSmallStyle.setFont(smallFont);

        CellStyle headerVerticalStyle = workbook.createCellStyle();
        headerVerticalStyle.cloneStyleFrom(headerStyle);
        headerVerticalStyle.setRotation((short) 90);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.setFont(titleFont);
        numberStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        numberStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(numberStyle);

        setMergedRegionWithStyle(sheet, 0, 0, 0, 22, titleStyle, "7.6 Результаты измерений шума");
        setMergedRegionWithStyle(sheet, 1, 1, 0, 22, titleStyle, "7.6.2. Шум:");

        Row headerTopRow = sheet.createRow(2);
        Row headerMiddleRow = sheet.createRow(3);
        Row headerBottomRow = sheet.createRow(4);
        headerTopRow.setHeightInPoints(20f);
        headerMiddleRow.setHeightInPoints(20f);
        headerBottomRow.setHeightInPoints(108f);

        setMergedRegionWithStyle(sheet, 2, 4, 0, 0, headerVerticalStyle, "№ п/п");
        setMergedRegionWithStyle(sheet, 2, 4, 1, 1, headerVerticalStyle, "№ точки измерения");
        setMergedRegionWithStyle(sheet, 2, 4, 2, 2, headerStyle, "Место измерений");
        setMergedRegionWithStyle(sheet, 2, 4, 3, 3, headerStyle,
                "Источник шума \n(тип, вид, марка, условия замера)");

        setMergedRegionWithStyle(sheet, 2, 2, 4, 8, headerStyle, "Характер шума");
        setMergedRegionWithStyle(sheet, 3, 3, 4, 5, headerSmallStyle, "по спектру");
        setMergedRegionWithStyle(sheet, 3, 3, 6, 8, headerSmallStyle,
                "по временным\nхарактеристикам");

        setCellValueWithStyle(sheet, 4, 4, "широкополосный", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 5, "тональный", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 6, "постоянный", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 7, "непостоянный", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 8, "импульсный", headerVerticalStyle);

        setMergedRegionWithStyle(sheet, 2, 3, 9, 17, headerStyle,
                "Уровни звукового давления (дБ) ± U (дБ) в октавных полосах частот " +
                        "со среднегеометрическими частотами (Гц)");

        setCellValueWithStyle(sheet, 4, 9, "31,5", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 10, "63", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 11, "125", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 12, "250", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 13, "500", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 14, "1000", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 15, "2000", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 16, "4000", headerVerticalStyle);
        setCellValueWithStyle(sheet, 4, 17, "8000", headerVerticalStyle);

        setMergedRegionWithStyle(sheet, 2, 4, 18, 18, headerVerticalStyle,
                "Уровни звука (дБА) \n±U (дБ)");
        setMergedRegionWithStyle(sheet, 2, 4, 19, 21, headerVerticalStyle,
                "Эквивалентные уровни звука,  (дБА) ±U (дБ)");
        setMergedRegionWithStyle(sheet, 2, 4, 22, 22, headerVerticalStyle,
                "Максимальные уровни звука  (дБА)");

        Row numberingRow = sheet.createRow(5);
        for (int col = 0; col <= 18; col++) {
            setCellValueWithStyle(numberingRow, col, String.valueOf(col + 1), numberStyle);
        }
        setMergedRegionWithStyle(sheet, 5, 5, 19, 21, numberStyle, "20");
        setCellValueWithStyle(numberingRow, 22, "21", numberStyle);
    }

    private static void addSimpleNumberingRow(Workbook workbook, Sheet sheet) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 9);

        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.setFont(titleFont);
        numberStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        numberStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        setThinBorders(numberStyle);

        Row numberingRow = sheet.createRow(0);
        for (int col = 0; col <= 18; col++) {
            setCellValueWithStyle(numberingRow, col, String.valueOf(col + 1), numberStyle);
        }
        setMergedRegionWithStyle(sheet, 0, 0, 19, 21, numberStyle, "20");
        setCellValueWithStyle(numberingRow, 22, "21", numberStyle);
    }

    private static void fillNoiseResultsFromProtocol(Sheet sourceSheet,
                                                     Workbook targetWorkbook,
                                                     Sheet targetSheet,
                                                     boolean hasFullHeader) {
        if (sourceSheet == null || targetSheet == null) {
            return;
        }
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = sourceSheet.getWorkbook()
                .getCreationHelper()
                .createFormulaEvaluator();
        CellStyle centerStyle = createNoiseDataStyle(targetWorkbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        CellStyle leftStyle = createNoiseDataStyle(targetWorkbook,
                org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

        int sourceRowIndex = 7;
        int targetRowIndex = hasFullHeader ? 6 : 1;
        int lastRow = sourceSheet.getLastRowNum();

        while (sourceRowIndex <= lastRow) {
            CellRangeAddress mergedRegion = findMergedRegion(sourceSheet, sourceRowIndex, 0);
            if (mergedRegion != null && mergedRegion.getFirstRow() == sourceRowIndex) {
                if (isMergedDateRow(mergedRegion)) {
                    String text = "Дата, время проведения измерений ____________________________";
                    setMergedRegionWithStyle(targetSheet, targetRowIndex, targetRowIndex,
                            0, 22, leftStyle, text);
                    adjustRowHeightForMergedText(targetSheet, targetRowIndex, 0, 22, text);
                    targetRowIndex++;
                    sourceRowIndex = mergedRegion.getLastRow() + 1;
                    continue;
                }
                if (isMergedBlockHeader(mergedRegion)) {
                    String aValue = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);
                    String bValue = readMergedCellValue(sourceSheet, sourceRowIndex, 1, formatter, evaluator);
                    String cValue = readMergedCellValue(sourceSheet, sourceRowIndex, 2, formatter, evaluator);
                    String dValue = readMergedCellValue(sourceSheet, sourceRowIndex, 3, formatter, evaluator);
                    targetRowIndex = appendNoiseMeasurementBlock(targetSheet, targetWorkbook, targetRowIndex,
                            aValue, bValue, cValue, dValue, centerStyle, leftStyle);
                    sourceRowIndex = mergedRegion.getLastRow() + 1;
                    continue;
                }
            }
            if (mergedRegion == null) {
                Row sourceRow = sourceSheet.getRow(sourceRowIndex);
                Cell cellA = sourceRow == null ? null : sourceRow.getCell(0);
                if (isNumericCell(cellA, formatter, evaluator)) {
                    CellRangeAddress dRegion = findMergedRegion(sourceSheet, sourceRowIndex, 3);
                    int blockStart = dRegion != null ? dRegion.getFirstRow() : sourceRowIndex;
                    int blockEnd = dRegion != null ? dRegion.getLastRow() : sourceRowIndex;
                    if (sourceRowIndex == blockStart) {
                        targetRowIndex = appendProtocolRows(sourceSheet, targetSheet, targetRowIndex,
                                blockStart, blockEnd, centerStyle, leftStyle, formatter, evaluator);
                        sourceRowIndex = blockEnd + 1;
                        continue;
                    }
                }
            }
            sourceRowIndex++;
        }
    }

    private static boolean isMergedDateRow(CellRangeAddress mergedRegion) {
        return mergedRegion.getFirstColumn() == 0
                && mergedRegion.getLastColumn() >= 23
                && mergedRegion.getFirstRow() == mergedRegion.getLastRow();
    }

    private static boolean isMergedBlockHeader(CellRangeAddress mergedRegion) {
        return mergedRegion.getFirstColumn() == 0
                && mergedRegion.getLastColumn() == 0
                && mergedRegion.getLastRow() - mergedRegion.getFirstRow() == 2;
    }

    private static int appendNoiseMeasurementBlock(Sheet sheet,
                                                   Workbook workbook,
                                                   int startRow,
                                                   String aValue,
                                                   String bValue,
                                                   String cValue,
                                                   String dValue,
                                                   CellStyle centerStyle,
                                                   CellStyle leftStyle) {
        int endRow = startRow + 4;
        setMergedRegionWithStyle(sheet, startRow, endRow, 0, 0, centerStyle, safe(aValue));
        setMergedRegionWithStyle(sheet, startRow, endRow, 1, 1, leftStyle, safe(bValue));

        setCellValueWithStyle(sheet, startRow, 2, safe(cValue), leftStyle);
        setCellValueWithStyle(sheet, startRow, 3, safe(dValue), leftStyle);
        fillRowCells(sheet, startRow, 4, 18, "", centerStyle);
        setMergedRegionWithStyle(sheet, startRow, startRow, 19, 21, centerStyle, "");
        setCellValueWithStyle(sheet, startRow, 22, "", centerStyle);

        int row2 = startRow + 1;
        setMergedRegionWithStyle(sheet, row2, row2, 2, 3, leftStyle, "Фон");
        fillRowCells(sheet, row2, 4, 18, "", centerStyle);
        setMergedRegionWithStyle(sheet, row2, row2, 19, 21, centerStyle, "");
        setCellValueWithStyle(sheet, row2, 22, "", centerStyle);

        int row3 = startRow + 2;
        String correctionSpectra = "Поправка (МИ Ш.13-2021 п.12.3.2.1.3) дБА (дБ) ";
        setMergedRegionWithStyle(sheet, row3, row3, 2, 8, leftStyle, correctionSpectra);
        fillRowCells(sheet, row3, 9, 18, "", centerStyle);
        setMergedRegionWithStyle(sheet, row3, row3, 19, 21, centerStyle, "");
        setCellValueWithStyle(sheet, row3, 22, "", centerStyle);
        adjustRowHeightForMergedText(sheet, row3, 2, 8, correctionSpectra);

        int row4 = startRow + 3;
        String correctionTime = "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ) ";
        setMergedRegionWithStyle(sheet, row4, row4, 2, 8, leftStyle, correctionTime);
        fillRowCells(sheet, row4, 9, 18, "2", centerStyle);
        setMergedRegionWithStyle(sheet, row4, row4, 19, 21, centerStyle, "2");
        setCellValueWithStyle(sheet, row4, 22, "2", centerStyle);
        adjustRowHeightForMergedText(sheet, row4, 2, 8, correctionTime);

        int row5 = startRow + 4;
        String levelsText = "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ) ";
        setMergedRegionWithStyle(sheet, row5, row5, 2, 8, leftStyle, levelsText);
        fillRowCells(sheet, row5, 9, 18, "", centerStyle);
        setMergedRegionWithStyle(sheet, row5, row5, 19, 21, centerStyle, "");
        setCellValueWithStyle(sheet, row5, 22, "", centerStyle);
        Row row5Ref = sheet.getRow(row5);
        if (row5Ref != null) {
            row5Ref.setHeightInPoints(sheet.getDefaultRowHeightInPoints());
        }

        return endRow + 1;
    }

    private static int appendProtocolRows(Sheet sourceSheet,
                                          Sheet targetSheet,
                                          int targetStartRow,
                                          int sourceStartRow,
                                          int sourceEndRow,
                                          CellStyle centerStyle,
                                          CellStyle leftStyle,
                                          DataFormatter formatter,
                                          FormulaEvaluator evaluator) {
        CellRangeAddress dRegion = findMergedRegion(sourceSheet, sourceStartRow, 3);
        String dValue = dRegion == null
                ? readCellText(sourceSheet.getRow(sourceStartRow), 3, formatter, evaluator)
                : readMergedCellValue(sourceSheet, sourceStartRow, 3, formatter, evaluator);
        int targetRowIndex = targetStartRow;
        for (int sourceRowIndex = sourceStartRow; sourceRowIndex <= sourceEndRow; sourceRowIndex++) {
            Row sourceRow = sourceSheet.getRow(sourceRowIndex);
            setCellValueWithStyle(targetSheet, targetRowIndex, 0,
                    readCellText(sourceRow, 0, formatter, evaluator), centerStyle);
            setCellValueWithStyle(targetSheet, targetRowIndex, 1,
                    readCellText(sourceRow, 1, formatter, evaluator), leftStyle);
            setCellValueWithStyle(targetSheet, targetRowIndex, 2,
                    readCellText(sourceRow, 2, formatter, evaluator), leftStyle);
            if (dRegion == null) {
                setCellValueWithStyle(targetSheet, targetRowIndex, 3,
                        readCellText(sourceRow, 3, formatter, evaluator), leftStyle);
            } else {
                String value = sourceRowIndex == sourceStartRow ? dValue : "";
                setCellValueWithStyle(targetSheet, targetRowIndex, 3, value, leftStyle);
            }
            fillRowCells(targetSheet, targetRowIndex, 4, 18, "", centerStyle);
            setMergedRegionWithStyle(targetSheet, targetRowIndex, targetRowIndex, 19, 21, centerStyle, "");
            setCellValueWithStyle(targetSheet, targetRowIndex, 22, "", centerStyle);
            targetRowIndex++;
        }
        if (dRegion != null && sourceEndRow > sourceStartRow) {
            setMergedRegionWithStyle(targetSheet, targetStartRow, targetRowIndex - 1, 3, 3, leftStyle, dValue);
        }
        return targetRowIndex;
    }

    private static void fillRowCells(Sheet sheet,
                                     int rowIndex,
                                     int startCol,
                                     int endCol,
                                     String value,
                                     CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        for (int col = startCol; col <= endCol; col++) {
            setCellValueWithStyle(row, col, value, style);
        }
    }

    private static CellStyle createNoiseDataStyle(Workbook workbook,
                                                  org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 9);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        style.setWrapText(true);
        setThinBorders(style);
        return style;
    }

    private static void setMergedRegionWithStyle(Sheet sheet,
                                                 int rowStart,
                                                 int rowEnd,
                                                 int colStart,
                                                 int colEnd,
                                                 CellStyle style,
                                                 String value) {
        Row row = sheet.getRow(rowStart);
        if (row == null) {
            row = sheet.createRow(rowStart);
        }
        Cell cell = row.getCell(colStart);
        if (cell == null) {
            cell = row.createCell(colStart);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);

        CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
        sheet.addMergedRegion(region);
        RegionUtil.setBorderTop(style.getBorderTop(), region, sheet);
        RegionUtil.setBorderBottom(style.getBorderBottom(), region, sheet);
        RegionUtil.setBorderLeft(style.getBorderLeft(), region, sheet);
        RegionUtil.setBorderRight(style.getBorderRight(), region, sheet);

        for (int rowIndex = rowStart; rowIndex <= rowEnd; rowIndex++) {
            Row currentRow = sheet.getRow(rowIndex);
            if (currentRow == null) {
                currentRow = sheet.createRow(rowIndex);
            }
            for (int colIndex = colStart; colIndex <= colEnd; colIndex++) {
                Cell currentCell = currentRow.getCell(colIndex);
                if (currentCell == null) {
                    currentCell = currentRow.createCell(colIndex);
                }
                currentCell.setCellStyle(style);
            }
        }
    }

    private static void setCellValueWithStyle(Sheet sheet, int rowIndex, int columnIndex,
                                              String value, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        setCellValueWithStyle(row, columnIndex, value, style);
    }

    private static void setCellValueWithStyle(Row row, int columnIndex, String value, CellStyle style) {
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static int addInstrumentRow(Sheet sheet,
                                        int rowIndex,
                                        String name,
                                        String serialNumber,
                                        String checkbox,
                                        CellStyle nameStyle,
                                        CellStyle serialStyle,
                                        CellStyle checkboxStyle) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        mergeCellRangeWithStyle(sheet, rowIndex, 1, 19, nameStyle, safe(name));
        mergeCellRangeWithStyle(sheet, rowIndex, 20, 29, serialStyle, safe(serialNumber));

        Cell checkboxCell = row.getCell(30);
        if (checkboxCell == null) {
            checkboxCell = row.createCell(30);
        }
        checkboxCell.setCellStyle(checkboxStyle);
        checkboxCell.setCellValue(safe(checkbox));

        adjustRowHeightForInstrumentRow(sheet, rowIndex, safe(name), safe(serialNumber));

        rowIndex++;
        return rowIndex;
    }

    private static void mergeCellRangeWithStyle(Sheet sheet,
                                                int rowIndex,
                                                int firstCol,
                                                int lastCol,
                                                CellStyle style,
                                                String value) {
        CellRangeAddress region = new CellRangeAddress(rowIndex, rowIndex, firstCol, lastCol);
        sheet.addMergedRegion(region);

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        for (int col = firstCol; col <= lastCol; col++) {
            Cell c = row.getCell(col);
            if (c == null) {
                c = row.createCell(col);
            }
            c.setCellStyle(style);
            if (col == firstCol) {
                c.setCellValue(value);
            }
        }

        RegionUtil.setBorderTop(style.getBorderTop(), region, sheet);
        RegionUtil.setBorderBottom(style.getBorderBottom(), region, sheet);
        RegionUtil.setBorderLeft(style.getBorderLeft(), region, sheet);
        RegionUtil.setBorderRight(style.getBorderRight(), region, sheet);
    }

    private static void setThinBorders(CellStyle style) {
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
    }

    private static void setMergedCellValue(Sheet sheet, int rowIndex, String text, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            cell = row.createCell(0);
        }
        cell.setCellStyle(style);
        cell.setCellValue(text);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 31));
    }

    private static void setMergedCellValueWithPrefix(Sheet sheet,
                                                     int rowIndex,
                                                     String prefix,
                                                     String value,
                                                     Font prefixFont,
                                                     Font valueFont,
                                                     CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(0);
        if (cell == null) {
            cell = row.createCell(0);
        }

        String text = (prefix == null ? "" : prefix) + safe(value);
        org.apache.poi.ss.usermodel.RichTextString richText =
                cell.getSheet().getWorkbook().getCreationHelper().createRichTextString(text);
        int prefixLength = prefix == null ? 0 : prefix.length();
        richText.applyFont(0, prefixLength, prefixFont);
        richText.applyFont(prefixLength, text.length(), valueFont);
        cell.setCellStyle(style);
        cell.setCellValue(richText);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 31));
    }

    private static int[] buildColumnWidthsPx() {
        int[] baseWidths = new int[32];
        baseWidths[0] = 36;
        baseWidths[1] = 33;
        baseWidths[2] = 32;
        baseWidths[3] = 32;
        baseWidths[4] = 32;
        baseWidths[5] = 32;
        baseWidths[6] = 32;
        baseWidths[7] = 32;
        baseWidths[8] = 32;
        for (int col = 9; col <= 17; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[18] = 53;
        baseWidths[19] = 41;
        baseWidths[20] = 58;
        baseWidths[21] = 32;
        for (int col = 22; col <= 29; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[30] = 36;
        baseWidths[31] = 11;

        int[] widths = new int[baseWidths.length];
        for (int i = 0; i < baseWidths.length; i++) {
            widths[i] = (int) Math.round(baseWidths[i] * COLUMN_WIDTH_SCALE);
        }
        return widths;
    }

    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static double cmToInches(double centimeters) {
        return centimeters / 2.54d;
    }

    private static float pixelsToPoints(int pixels) {
        return pixels * 0.75f;
    }

    private static float pointsToPixels(float points) {
        return points / 0.75f;
    }

    private static MapHeaderData resolveHeaderData(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return new MapHeaderData("", "", "", "", "", "");
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return new MapHeaderData("", "", "", "", "", "");
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findHeaderData(sheet);
        } catch (Exception ex) {
            return new MapHeaderData("", "", "", "", "", "");
        }
    }

    private static MapHeaderData findHeaderData(Sheet sheet) {
        String customer = "";
        String dates = "";
        String representative = "";
        String legalAddress = "";
        String objectName = "";
        String objectAddress = "";
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell);
                String text = rawText.trim();
                String normalized = normalizeText(rawText);
                if (customer.isEmpty()) {
                    customer = extractCustomer(text);
                    if (customer.isEmpty() && text.startsWith(CUSTOMER_PREFIX)) {
                        customer = readNextCellText(row, cell, formatter);
                    }
                }
                if (dates.isEmpty()) {
                    dates = extractMeasurementDates(text);
                }
                if (representative.isEmpty()) {
                    representative = extractRepresentative(text);
                    if (representative.isEmpty() && text.startsWith(REPRESENTATIVE_PREFIX)) {
                        representative = readNextCellText(row, cell, formatter);
                    }
                }
                if (legalAddress.isEmpty()) {
                    legalAddress = extractLegalAddress(normalized);
                    if (legalAddress.isEmpty() && normalized.startsWith(LEGAL_ADDRESS_PREFIX)) {
                        legalAddress = readNextCellText(row, cell, formatter);
                    }
                }
                if (objectName.isEmpty()) {
                    objectName = extractObjectName(normalized);
                    if (objectName.isEmpty() && normalized.startsWith(OBJECT_NAME_PREFIX)) {
                        objectName = readNextCellText(row, cell, formatter);
                    }
                }
                if (objectAddress.isEmpty()) {
                    objectAddress = extractObjectAddress(normalized);
                    if (objectAddress.isEmpty() && normalized.startsWith(OBJECT_ADDRESS_PREFIX)) {
                        objectAddress = readNextCellText(row, cell, formatter);
                    }
                }
                if (!customer.isEmpty() && !dates.isEmpty() && !representative.isEmpty()
                        && !legalAddress.isEmpty() && !objectName.isEmpty() && !objectAddress.isEmpty()) {
                    return new MapHeaderData(customer, dates, representative, legalAddress, objectName, objectAddress);
                }
            }
        }
        return new MapHeaderData(customer, dates, representative, legalAddress, objectName, objectAddress);
    }

    private static String extractCustomer(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(CUSTOMER_PREFIX);
        if (index < 0) {
            return "";
        }
        String tail = text.substring(index + CUSTOMER_PREFIX.length()).trim();
        return tail;
    }

    private static String extractMeasurementDates(String text) {
        if (text == null || text.isEmpty()) {
            return "";
        }
        int start = text.indexOf(MEASUREMENT_DATES_PHRASE);
        if (start < 0) {
            return "";
        }
        String tail = text.substring(start + MEASUREMENT_DATES_PHRASE.length()).trim();
        if (tail.isEmpty()) {
            return "";
        }
        java.util.regex.Matcher matcher = DATE_PATTERN.matcher(tail);
        if (!matcher.find()) {
            return tail;
        }
        java.util.List<String> dates = new java.util.ArrayList<>();
        matcher.reset();
        while (matcher.find()) {
            dates.add(matcher.group());
        }
        return String.join(", ", dates);
    }

    private static String extractRepresentative(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(REPRESENTATIVE_PREFIX);
        if (index < 0) {
            return "";
        }
        String tail = text.substring(index + REPRESENTATIVE_PREFIX.length()).trim();
        return tail;
    }

    private static String extractLegalAddress(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(LEGAL_ADDRESS_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + LEGAL_ADDRESS_PREFIX.length()).trim();
    }

    private static String extractObjectName(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(OBJECT_NAME_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + OBJECT_NAME_PREFIX.length()).trim();
    }

    private static String extractObjectAddress(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(OBJECT_ADDRESS_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + OBJECT_ADDRESS_PREFIX.length()).trim();
    }

    private static String readNextCellText(Row row, Cell cell, DataFormatter formatter) {
        if (row == null || cell == null) {
            return "";
        }
        Cell next = row.getCell(cell.getColumnIndex() + 1);
        if (next == null) {
            return "";
        }
        String nextText = formatter.formatCellValue(next).trim();
        return nextText;
    }

    private static String resolveMeasurementPerformer(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            for (int idx = workbook.getNumberOfSheets() - 1; idx >= 0; idx--) {
                if (idx == 0) {
                    continue;
                }
                Sheet sheet = workbook.getSheetAt(idx);
                if (sheet == null) {
                    continue;
                }
                if (isGeneratorSheet(sheet.getSheetName())) {
                    continue;
                }
                String performer = findMeasurementPerformer(sheet, formatter);
                if (!performer.isEmpty()) {
                    return performer;
                }
            }
            return "";
        } catch (Exception ex) {
            return "";
        }
    }

    private static String resolveControlDate(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findControlDate(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String resolveTitleMeasurementDates(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findTitleMeasurementDates(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findTitleMeasurementDates(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        Row row = sheet.getRow(TITLE_MEASUREMENT_DATES_ROW);
        if (row == null) {
            return "";
        }
        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (text.isEmpty()) {
                continue;
            }
            String prefix = "2. Дата замеров:";
            int index = text.indexOf(prefix);
            if (index >= 0) {
                return text.substring(index + prefix.length()).trim();
            }
            return text;
        }
        return "";
    }

    private static String resolveMeasurementDates(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() <= 1) {
                return "";
            }
            DataFormatter formatter = new DataFormatter();
            java.util.Map<String, String> dates = new java.util.LinkedHashMap<>();
            for (int idx = 1; idx < workbook.getNumberOfSheets(); idx++) {
                Sheet sheet = workbook.getSheetAt(idx);
                if (sheet == null) {
                    continue;
                }
                if (isGeneratorSheet(sheet.getSheetName())) {
                    continue;
                }
                collectMeasurementDates(sheet, formatter, dates);
            }
            return String.join(", ", dates.values());
        } catch (Exception ex) {
            return "";
        }
    }

    private static java.util.List<String> extractMeasurementDatesList(String datesText) {
        if (datesText == null || datesText.isBlank()) {
            return java.util.List.of("");
        }
        java.util.LinkedHashSet<String> dates = new java.util.LinkedHashSet<>();
        java.util.regex.Matcher matcher = DATE_PATTERN.matcher(datesText);
        while (matcher.find()) {
            dates.add(matcher.group());
        }
        if (dates.isEmpty()) {
            return java.util.List.of(datesText.trim());
        }
        return new java.util.ArrayList<>(dates);
    }

    private static String findControlDate(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        Row row = sheet.getRow(6);
        if (row == null) {
            return "";
        }
        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (text.isEmpty()) {
                continue;
            }
            return text;
        }
        return "";
    }

    private static String resolveSpecialConditions(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            boolean hasRadonSheet = false;
            boolean hasArtificialLightingSheet = false;
            for (int idx = 0; idx < workbook.getNumberOfSheets(); idx++) {
                Sheet sheet = workbook.getSheetAt(idx);
                if (sheet == null) {
                    continue;
                }
                String normalized = normalizeText(sheet.getSheetName()).toLowerCase(Locale.ROOT);
                if (normalized.equals("эроа радона")) {
                    hasRadonSheet = true;
                }
                if (normalized.equals("иск освещение")) {
                    hasArtificialLightingSheet = true;
                }
            }
            java.util.List<String> conditions = new java.util.ArrayList<>();
            if (hasRadonSheet) {
                conditions.add("заказчик сообщил, что перед измерением ЭРОА радона "
                        + "здание выдерженно в течении более 12 часов при закрытых дверях и окнах");
            }
            if (hasArtificialLightingSheet) {
                conditions.add("Отношение естественной освещенности к искусственной составляет");
            }
            return String.join(". ", conditions);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String resolveAdditionalInfo(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findAdditionalInfo(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findAdditionalInfo(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = normalizeText(formatter.formatCellValue(cell));
                String lower = text.toLowerCase(Locale.ROOT);
                int markerIndex = lower.indexOf(ADDITIONAL_INFO_PREFIX);
                if (markerIndex < 0) {
                    continue;
                }
                String tail = text.substring(markerIndex + ADDITIONAL_INFO_PREFIX.length()).trim();
                return tail;
            }
        }
        return "";
    }

    private static String resolveMeasurementMethods(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findMeasurementMethods(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static java.util.List<InstrumentData> resolveMeasurementInstruments(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return java.util.Collections.emptyList();
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return java.util.Collections.emptyList();
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findMeasurementInstruments(sheet);
        } catch (Exception ex) {
            return java.util.Collections.emptyList();
        }
    }

    private static java.util.List<InstrumentData> findMeasurementInstruments(Sheet sheet) {
        if (sheet == null) {
            return java.util.Collections.emptyList();
        }
        DataFormatter formatter = new DataFormatter();
        int sectionRowIndex = -1;
        int headerRowIndex = -1;
        int nameColumn = -1;
        int serialColumn = -1;
        for (Row row : sheet) {
            for (Cell cell : row) {
                String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                if (normalized.equals(INSTRUMENTS_SECTION_HEADER)) {
                    if (sectionRowIndex < 0) {
                        sectionRowIndex = row.getRowNum();
                    }
                }
            }
        }

        if (sectionRowIndex >= 0) {
            for (int rowIndex = sectionRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }
                for (Cell cell : row) {
                    String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                    if (normalized.equals(INSTRUMENTS_NAME_HEADER)) {
                        nameColumn = cell.getColumnIndex();
                        headerRowIndex = rowIndex;
                        break;
                    }
                }
                if (headerRowIndex >= 0) {
                    break;
                }
            }
        }

        if (headerRowIndex < 0 || nameColumn < 0) {
            return java.util.Collections.emptyList();
        }

        Row headerRow = sheet.getRow(headerRowIndex);
        if (headerRow != null) {
            for (Cell cell : headerRow) {
                String normalized = normalizeText(formatter.formatCellValue(cell)).toLowerCase(Locale.ROOT);
                if (normalized.contains("завод")) {
                    serialColumn = cell.getColumnIndex();
                    break;
                }
            }
        }

        if (serialColumn < 0) {
            serialColumn = nameColumn + 19;
        }

        java.util.List<InstrumentData> instruments = new java.util.ArrayList<>();
        java.util.Set<String> seenSerials = new java.util.HashSet<>();
        FormulaEvaluator evaluator = sheet.getWorkbook()
                .getCreationHelper()
                .createFormulaEvaluator();
        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            String name = readMergedCellValue(sheet, rowIndex, nameColumn, formatter, evaluator);
            String serial = readMergedCellValue(sheet, rowIndex, serialColumn, formatter, evaluator);
            String normalizedName = normalizeText(name).toLowerCase(Locale.ROOT);
            if (normalizedName.equals(INSTRUMENTS_NAME_HEADER)) {
                continue;
            }
            if (name.isBlank() && serial.isBlank()) {
                break;
            }
            String serialKey = normalizeText(serial).toLowerCase(Locale.ROOT);
            if (!serialKey.isBlank() && !seenSerials.add(serialKey)) {
                continue;
            }
            instruments.add(new InstrumentData(name, serial));
        }
        return instruments;
    }

    private static String readMergedCellValue(Sheet sheet,
                                              int rowIndex,
                                              int colIndex,
                                              DataFormatter formatter,
                                              FormulaEvaluator evaluator) {
        if (sheet == null) {
            return "";
        }
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(rowIndex, colIndex)) {
                Row row = sheet.getRow(range.getFirstRow());
                if (row == null) {
                    return "";
                }
                Cell cell = row.getCell(range.getFirstColumn());
                return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
            }
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
    }

    private static String readCellText(Row row,
                                       int columnIndex,
                                       DataFormatter formatter,
                                       FormulaEvaluator evaluator) {
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(columnIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
    }

    private static boolean isNumericCell(Cell cell,
                                         DataFormatter formatter,
                                         FormulaEvaluator evaluator) {
        if (cell == null) {
            return false;
        }
        String text = normalizeText(formatter.formatCellValue(cell, evaluator));
        if (text.isBlank()) {
            return false;
        }
        String normalized = text.replace(',', '.');
        try {
            Double.parseDouble(normalized);
            return true;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) {
            return null;
        }
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String findMeasurementMethods(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell);
                String normalized = normalizeText(rawText);
                if (!normalized.equals(METHODS_HEADER)) {
                    continue;
                }
                return collectMethodsBelow(sheet, row.getRowNum(), cell.getColumnIndex(), formatter);
            }
        }
        return "";
    }

    private static String collectMethodsBelow(Sheet sheet, int headerRow, int columnIndex, DataFormatter formatter) {
        java.util.List<String> methods = new java.util.ArrayList<>();
        for (int rowIndex = headerRow + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                break;
            }
            Cell cell = row.getCell(columnIndex);
            String raw = cell == null ? "" : formatter.formatCellValue(cell);
            String normalized = normalizeText(raw);
            if (normalized.isEmpty()) {
                break;
            }
            if (normalized.equals(METHODS_HEADER)) {
                continue;
            }
            methods.add(normalized);
        }
        return String.join("; ", methods);
    }

    private static boolean isGeneratorSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.equals("генератор")
                || normalized.equals("генератор (2)")
                || normalized.equals("генератор (2.0.)");
    }

    private static boolean isNoiseProtocolSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.contains("шум");
    }

    private static boolean isMicroclimateSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.equals("микроклимат");
    }

    private static boolean isIgnoredProtocolSheet(String sheetName) {
        if (sheetName == null) {
            return false;
        }
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.equals("неопред");
    }

    private static String findMeasurementPerformer(Sheet sheet, DataFormatter formatter) {
        if (sheet == null) {
            return "";
        }
        for (Row row : sheet) {
            boolean hasProtocolPrepared = false;
            for (Cell cell : row) {
                String text = normalizeText(formatter.formatCellValue(cell));
                if (text.contains("Протокол подготовил")) {
                    hasProtocolPrepared = true;
                    break;
                }
            }
            if (!hasProtocolPrepared) {
                continue;
            }
            String performer = resolvePerformerFromRow(row, formatter);
            if (!performer.isEmpty()) {
                return performer;
            }
        }
        return "";
    }

    private static String resolvePerformerFromRow(Row row, DataFormatter formatter) {
        if (row == null) {
            return "";
        }
        String rowText = collectRowText(row, formatter).toLowerCase(Locale.ROOT);
        if (rowText.contains("белов")) {
            return "Белов Д.А.";
        }
        if (rowText.contains("тарновский")) {
            return "Тарновский М.О.";
        }
        return "";
    }

    private static String collectRowText(Row row, DataFormatter formatter) {
        StringBuilder builder = new StringBuilder();
        for (Cell cell : row) {
            String text = normalizeText(formatter.formatCellValue(cell));
            if (!text.isEmpty()) {
                if (builder.length() > 0) {
                    builder.append(' ');
                }
                builder.append(text);
            }
        }
        return builder.toString();
    }

    private static void collectMeasurementDates(Sheet sheet,
                                                DataFormatter formatter,
                                                java.util.Map<String, String> dates) {
        if (sheet == null) {
            return;
        }
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            // Ищем даты в объединениях A:Y (включительно), где обычно лежит строка
            // "Дата, время проведения измерений ..." в шаблонах для шумов.
            if (range.getFirstColumn() > 24 || range.getLastColumn() > 24) {
                continue;
            }
            Row row = sheet.getRow(range.getFirstRow());
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(range.getFirstColumn());
            String text = normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
            if (text.isEmpty()) {
                continue;
            }
            java.util.regex.Matcher matcher = MEASUREMENT_DATE_PATTERN.matcher(text);
            while (matcher.find()) {
                String rawDate = matcher.group();
                String key = normalizeText(rawDate).toLowerCase(Locale.ROOT);
                dates.putIfAbsent(key, rawDate);
            }
        }
    }

    private static String resolveProtocolNumber(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findProtocolNumber(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findProtocolNumber(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                int index = text.indexOf(PROTOCOL_PREFIX);
                if (index < 0) {
                    continue;
                }
                String tail = text.substring(index + PROTOCOL_PREFIX.length()).trim();
                tail = stripLeadingNumberMarker(tail);
                if (!tail.isEmpty()) {
                    return tail;
                }
                String nextText = readNextCellText(row, cell, formatter);
                return stripLeadingNumberMarker(nextText);
            }
        }
        return "";
    }

    private static String resolveContractText(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findContractText(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findContractText(Sheet sheet) {
        if (sheet == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                int index = text.indexOf(BASIS_PREFIX);
                if (index < 0) {
                    continue;
                }
                String tail = text.substring(index + BASIS_PREFIX.length()).trim();
                if (!tail.isEmpty()) {
                    return tail;
                }
                return readNextCellText(row, cell, formatter);
            }
        }
        return "";
    }

    private static String normalizeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }

    private static String resolveControlSuffix(String measurementPerformer) {
        if ("Тарновский М.О.".equals(measurementPerformer)) {
            return "Инженер Белов Д.А.";
        }
        if ("Белов Д.А.".equals(measurementPerformer)) {
            return "Заведующий лабораторией Тарновский М.О.";
        }
        return "";
    }

    private static String safe(String value) {
        return value == null ? "" : value.trim();
    }

    private static String stripLeadingNumberMarker(String value) {
        if (value == null) {
            return "";
        }
        String trimmed = value.trim();
        if (trimmed.startsWith("№")) {
            return trimmed.substring(1).trim();
        }
        return trimmed;
    }

    private static void adjustRowHeightForMergedTextDoubling(Sheet sheet,
                                                             int rowIndex,
                                                             int firstCol,
                                                             int lastCol,
                                                             String text) {
        if (sheet == null) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double totalChars = totalColumnChars(sheet, firstCol, lastCol);
        int lines = estimateWrappedLines(text, totalChars);

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        int multiplier = lines > 1 ? 2 : 1;

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * multiplier)));
    }

    private static void adjustRowHeightForMergedText(Sheet sheet,
                                                     int rowIndex,
                                                     int firstCol,
                                                     int lastCol,
                                                     String text) {
        if (sheet == null) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double totalChars = totalColumnChars(sheet, firstCol, lastCol);
        int lines = estimateWrappedLines(text, totalChars);

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * Math.max(1, lines))));
    }

    private static void adjustRowHeightForInstrumentRow(Sheet sheet,
                                                        int rowIndex,
                                                        String name,
                                                        String serialNumber) {
        if (sheet == null) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double nameChars = totalColumnChars(sheet, 1, 19);
        double serialChars = totalColumnChars(sheet, 20, 29);
        int nameLines = estimateWrappedLines(name, nameChars);
        int serialLines = estimateWrappedLines(serialNumber, serialChars);
        int lines = Math.max(1, Math.max(nameLines, serialLines));

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        row.setHeightInPoints(pixelsToPoints((int) (baseHeightPx * lines)));
    }

    private static void addRowHeightPixels(Sheet sheet, int rowIndex, int extraPixels) {
        if (sheet == null || extraPixels <= 0) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        float currentHeightPx = pointsToPixels(row.getHeightInPoints());
        if (currentHeightPx <= 0f) {
            currentHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        row.setHeightInPoints(pixelsToPoints((int) (currentHeightPx + extraPixels)));
    }

    private static double totalColumnChars(Sheet sheet, int firstCol, int lastCol) {
        double totalChars = 0.0;
        for (int c = firstCol; c <= lastCol; c++) {
            totalChars += sheet.getColumnWidth(c) / 256.0;
        }
        return Math.max(1.0, totalChars);
    }

    private static int estimateWrappedLines(String text, double colChars) {
        if (text == null || text.isBlank()) {
            return 1;
        }

        int lines = 0;
        String[] segments = text.split("\\r?\\n");
        for (String seg : segments) {
            int len = Math.max(1, seg.trim().length());
            lines += (int) Math.ceil(len / Math.max(1.0, colChars));
        }
        return Math.max(1, lines);
    }

    private static final String CUSTOMER_PREFIX =
            "Наименование и контактные данные заявителя (заказчика):";
    private static final String MEASUREMENT_DATES_PHRASE = "Измерения были проведены";
    private static final String REPRESENTATIVE_PREFIX =
            "Измерения проводились в присутствии представителя заказчика:";
    private static final String LEGAL_ADDRESS_PREFIX = "Юридический адрес заказчика:";
    private static final String OBJECT_NAME_PREFIX =
            "Наименование предприятия, организации, объекта, где производились измерения:";
    private static final String OBJECT_ADDRESS_PREFIX = "Адрес предприятия (объекта):";
    private static final String METHODS_HEADER =
            "Документы, устанавливающие правила и методы исследований (испытаний) и измерений";
    private static final String ADDITIONAL_INFO_PREFIX =
            "дополнительные сведения (характеристика объекта):";
    private static final String INSTRUMENTS_SECTION_HEADER = "сведения о средствах измерения:";
    private static final String INSTRUMENTS_NAME_HEADER = "наименование, тип средства измерения";
    private static final String PROTOCOL_PREFIX = "Протокол испытаний";
    private static final String BASIS_PREFIX = "Основание для измерений: договор";
    private static final java.util.regex.Pattern DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{2}\\.\\d{2}\\.\\d{4}\\b");
    private static final java.util.regex.Pattern CONTROL_DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{1,2}\\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s+\\d{4}\\b",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);
    private static final java.util.regex.Pattern MEASUREMENT_DATE_PATTERN =
            java.util.regex.Pattern.compile(
                    "\\b\\d{2}\\.\\d{2}\\.(?:\\d{2}|\\d{4})\\b|\\b\\d{1,2}\\s+"
                            + "(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)"
                            + "\\s+\\d{4}\\b",
                    java.util.regex.Pattern.CASE_INSENSITIVE | java.util.regex.Pattern.UNICODE_CASE);

    private static final class MapHeaderData {
        private final String customerNameAndContacts;
        private final String measurementDates;
        private final String representative;
        private final String customerLegalAddress;
        private final String objectName;
        private final String objectAddress;

        private MapHeaderData(String customerNameAndContacts,
                              String measurementDates,
                              String representative,
                              String customerLegalAddress,
                              String objectName,
                              String objectAddress) {
            this.customerNameAndContacts = customerNameAndContacts;
            this.measurementDates = measurementDates;
            this.representative = representative;
            this.customerLegalAddress = customerLegalAddress;
            this.objectName = objectName;
            this.objectAddress = objectAddress;
        }
    }

    private static final class InstrumentData {
        private final String name;
        private final String serialNumber;

        private InstrumentData(String name, String serialNumber) {
            this.name = safe(name);
            this.serialNumber = safe(serialNumber);
        }
    }
}
