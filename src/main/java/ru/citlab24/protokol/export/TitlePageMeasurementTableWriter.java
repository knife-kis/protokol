package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import static ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon.estimateWrappedLines;

final class TitlePageMeasurementTableWriter {
    private TitlePageMeasurementTableWriter() {
    }

    static int[][] headerRanges() {
        return new int[][]{
                {0, 3},
                {4, 9},
                {10, 12},
                {13, 19},
                {20, 25}
        };
    }

    static String[] headerTexts() {
        return new String[]{
                "Измеряемый показатель",
                "Наименование, тип средства Измерения",
                "Заводской номер",
                "Погрешность средства измерения",
                "Сведения о поверке (№ свидетельства, срок действия)"
        };
    }

    static void rebuildTable(Workbook wb, String additionalInfoText) {
        if (wb == null) return;

        Sheet sheet = wb.getSheet("Титульная страница");
        if (sheet == null) return;

        // стили делаем заново (чтобы не зависеть от локальных переменных appendTitleSheet)
        Font baseFont = wb.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        Font smallFont = wb.createFont();
        smallFont.setFontName("Arial");
        smallFont.setFontHeightInPoints((short) 8);

        CellStyle headerStyle = wb.createCellStyle();
        headerStyle.setFont(baseFont);
        headerStyle.setWrapText(true);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        setThinBorders(headerStyle);

        CellStyle measurementStyle = wb.createCellStyle();
        measurementStyle.cloneStyleFrom(headerStyle);
        measurementStyle.setFont(smallFont);
        setThinBorders(measurementStyle);

        int headerRow = 35;

        // перезаписываем таблицу (setMergedText всё равно проставит текст/стили заново)
        write(sheet, headerRow, headerStyle, measurementStyle, additionalInfoText);

        // подгоняем высоту строки заголовка таблицы под переносы
        // (это твоя логика из AllExcelExporter)
        adjustHeaderRowHeightForMergedSections(
                sheet,
                headerRow,
                headerRanges(),
                headerTexts()
        );
    }

    private static void setThinBorders(CellStyle style) {
        if (style == null) return;
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    // локальный аналог твоего adjustRowHeightForMergedSections для заголовка таблицы
    private static void adjustHeaderRowHeightForMergedSections(Sheet sheet,
                                                               int rowIndex,
                                                               int[][] ranges,
                                                               String[] texts) {
        if (sheet == null || ranges == null || texts == null || ranges.length != texts.length) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);

        float baseHeightPoints = row.getHeightInPoints();
        if (baseHeightPoints <= 0) baseHeightPoints = 15f;

        int maxLines = 1;
        for (int i = 0; i < ranges.length; i++) {
            int[] range = ranges[i];
            if (range == null || range.length != 2) continue;

            double totalChars = 0.0;
            for (int c = range[0]; c <= range[1]; c++) {
                totalChars += sheet.getColumnWidth(c) / 256.0;
            }
            totalChars = Math.max(1.0, totalChars);

            int lines = estimateWrappedLines(texts[i], totalChars);
            maxLines = Math.max(maxLines, lines);
        }

        row.setHeightInPoints(baseHeightPoints * maxLines);
    }


    static void write(Sheet sheet, int headerRow, CellStyle headerStyle, CellStyle measurementStyle, String additionalInfoText) {
        String[] headers = headerTexts();
        writeHeaderRow(sheet, headerRow, headerStyle, headers);

        String indicatorsText = readIndicatorsText(sheet).toLowerCase();
        boolean hasResultantTemperature = indicatorsText.contains("результирующая температура")
                && hasResultantTemperatureValues(sheet.getWorkbook());
        boolean hasPulsation = indicatorsText.contains("коэффициент пульсации");
        boolean hasLightingIndicator = indicatorsText.contains("освещенность (");
        boolean hasKeoIndicator = indicatorsText.contains("коэффициент естественной освещенности");
        boolean hasStreetLightingIndicator =
                indicatorsText.contains("средняя горизонтальная освещенность на уровне земли");

        boolean hasMedSheet = sheet.getWorkbook().getSheet("МЭД") != null
                || sheet.getWorkbook().getSheet("МЭД (2)") != null;
        boolean hasEroaSheet = sheet.getWorkbook().getSheet("ЭРОА радона") != null;
        boolean hasLightingSheets = sheet.getWorkbook().getSheet("Иск освещение") != null
                || sheet.getWorkbook().getSheet("Иск освещение (2)") != null
                || sheet.getWorkbook().getSheet("КЕО") != null;
        boolean hasLightingPowerSheet = sheet.getWorkbook().getSheet("Иск освещение") != null
                || sheet.getWorkbook().getSheet("Иск освещение (2)") != null;
        boolean hasVentilationSheet = hasSheetWithPrefix(sheet.getWorkbook(), "Вентиляция");

        int rowIndex = headerRow + 1;
        setRowHeightPx(sheet, rowIndex, 126f);
        writeMeasurementRow(sheet, measurementStyle, rowIndex,
                "Длительность интервала времени",
                "Секундомеры электронные, Интеграл С-01",
                "\u00b1(9,6\u00b710-6 \u00b7\u0422x+0,01) \u0441\n" +
                        "\u0414\u043e\u043f\u043e\u043b\u043d\u0438\u0442\u0435\u043b\u044c\u043d\u0430\u044f \u0430\u0431\u0441\u043e\u043b\u044e\u0442\u043d\u0430\u044f \u043f\u043e\u0433\u0440\u0435\u0448\u043d\u043e\u0441\u0442\u044c \u043f\u0440\u0438 \u043e\u0442\u043a\u043b\u043e\u043d\u0435\u043d\u0438\u0438 \u0442\u0435\u043c\u043f\u0435\u0440\u0430\u0442\u0443\u0440\u044b \u043e\u0442 \u043d\u043e\u0440\u043c\u0430\u043b\u044c\u043d\u044b\u0445 \u0443\u0441\u043b\u043e\u0432\u0438\u0439 25 \u00b1 5 (\u02da\u0421) \u043d\u0430 1 \u02da\u0421 \u0438\u0437\u043c\u0435\u043d\u0435\u043d\u0438\u044f \u0442\u0435\u043c\u043f\u0435\u0440\u0430\u0442\u0443\u0440\u044b:\n" +
                        "-(2,2\u00b710-6\u00b7\u0422x) \u0441",
                "\u0421-\u0410\u0428/21-05-2025/433383424 \u0434\u043e 20.05.2026",
                "462667");
        rowIndex++;

        int microStartRow = rowIndex;
        String instrumentValue = "Комплект измерительный параметров микроклимата «Метеоскоп-М»";
        String verificationValue = "С-А/22-01-2024/310286448 до 21.01.2026";
        setMergedText(sheet, measurementStyle, microStartRow,
                microStartRow + (hasResultantTemperature ? 4 : 3), 4, 9, instrumentValue);
        setMergedText(sheet, measurementStyle, microStartRow,
                microStartRow + (hasResultantTemperature ? 4 : 3), 10, 12, "569521");
        setMergedText(sheet, measurementStyle, microStartRow,
                microStartRow + (hasResultantTemperature ? 4 : 3), 20, 25, verificationValue);

        setRowHeightPx(sheet, rowIndex, 33f);
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 0, 3, "Температура окружающего воздуха");
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 13, 19, "\u00b10,2 \u00b0C");
        rowIndex++;

        if (hasResultantTemperature) {
            setRowHeightPx(sheet, rowIndex, 29f);
            setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 0, 3, "Результирующая температура");
            setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 13, 19, "\u00b10,5 \u00b0C");
            rowIndex++;
        }

        setRowHeightPx(sheet, rowIndex, 30f);
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 0, 3, "Относительная влажность");
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 13, 19, "\u00b1 3%");
        rowIndex++;

        setRowHeightPx(sheet, rowIndex, 72f);
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 0, 3,
                "Скорость воздушного потока, скорость движения воздуха");
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 13, 19,
                "\u00b1(0,05+0,05V) м/с в диапазоне от 0,1 до 1 м/с включ.\n" +
                        "\u00b1(0,1+0,05V) м/с в диапазоне св. 1 до 20 м/с.\n" +
                        "где V – значение измеряемой скорости, м/с");
        rowIndex++;

        setRowHeightPx(sheet, rowIndex, 37f);
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 0, 3, "Атмосферное давление");
        setMergedText(sheet, measurementStyle, rowIndex, rowIndex, 13, 19, "\u00b1 0,13 кПа\n(\u00b11 мм рт.ст)");
        rowIndex++;

        int distanceRow = rowIndex;
        setRowHeightPx(sheet, distanceRow, 35f);
        writeMeasurementRow(sheet, measurementStyle, distanceRow,
                "Измерение расстояний",
                "Дальномер лазерный ADA Cosmo 70",
                "\u00b11,5 мм",
                "С-АШ/22-01-2025/403727315  до 21.01.2026",
                "000873");

        int medRow = distanceRow + 1;
        setRowHeightPx(sheet, medRow, 45f);
        if (hasMedSheet) {
            writeMeasurementRow(sheet, measurementStyle, medRow,
                    "Мощность дозы гамма-излучения",
                    "Дозиметр радиометр ДРБП-03",
                    "\u00b1 (15+4/Н)%\nГде Н – измеренные числовые значения МЭД",
                    "С-АШ/16-01-2025/402566181 до 15.01.2026",
                    "21115");
        } else {
            writeMeasurementRow(sheet, measurementStyle, medRow, "", "", "", "", "");
        }

        int eroaRow = medRow + 1;
        setRowHeightPx(sheet, eroaRow, 166f);
        if (hasEroaSheet) {
            writeMeasurementRow(sheet, measurementStyle, eroaRow,
                    "Эквивалентная равновесная объемная активность (ЭРОА) радона; " +
                            "Эквивалентная равновесная объемная активность (ЭРОА) торона",
                    "Аэрозольный альфа-радиометр РАА-20П2 «Поиск»",
                    "\u00b1 30%",
                    "С-ТТ/27-01-2025/405177537 до 26.01.2026",
                    "412");
        } else {
            writeMeasurementRow(sheet, measurementStyle, eroaRow, "", "", "", "", "");
        }

        int headerRepeatRow = eroaRow + 1;
        float headerHeight = getRowHeightPoints(sheet, headerRow);
        if (hasResultantTemperature) {
            setRowHeightPx(sheet, headerRepeatRow, 49f);
        } else if (headerHeight > 0) {
            setRowHeightPoints(sheet, headerRepeatRow, headerHeight);
        }
        writeHeaderRow(sheet, headerRepeatRow, headerStyle, headers);

        int lightingRow = headerRepeatRow + 1;
        setRowHeightPx(sheet, lightingRow, 125f);
        if (hasLightingSheets) {
            String lightingIndicators = buildLightingIndicatorText(
                    hasLightingIndicator,
                    hasKeoIndicator,
                    hasPulsation,
                    hasStreetLightingIndicator);
            String lightingErrorText = "\u00b1 8% (Освещенность)";
            if (hasPulsation) {
                lightingErrorText += " / \u00b1 10% (Коэффициент пульсации)";
            }
            writeMeasurementRow(sheet, measurementStyle, lightingRow,
                    lightingIndicators,
                    "Прибор комбинированный еЕлайт «еЛайт 01» исполнение 1, " +
                            "в комплектность входит: фотометрическая головка «еЛайт 03»; " +
                            "блок отображения информации БОИ-01 ",
                    lightingErrorText,
                    "С-ВЬ/30-10-2025/477698253 до 29.10.2027",
                    "Фотометрическая головка «еЛайт 03» зав. №03810-23; " +
                            "блок отображения информации БОИ-01 зав. №01807-23 Инв. № 31");
        } else {
            writeMeasurementRow(sheet, measurementStyle, lightingRow, "", "", "", "", "");
        }

        int multimeterRow = lightingRow + 1;
        setRowHeightPx(sheet, multimeterRow, 120f);
        if (hasLightingPowerSheet) {
            writeMeasurementRow(sheet, measurementStyle, multimeterRow,
                    "Измерение напряжения и частоты переменного ток",
                    "Мультиметр\nЦифровой, АРPА-103N",
                    "\u00b1(0,013×X + 5×к) для 40 Гц…1 кГц\n" +
                            "\u00b1(0,015×X + 5×к) в полосе частот 50…60 Гц\n" +
                            "\u00b1(0,0001*X + 1*к) \n" +
                            "X – значение измеренной величины по встроенному индикатору, к – цена единицы младшего разряда индикатора",
                    "С-АШ/17-01-2025/402863148 до 16.01.2026",
                    "05151010");
        } else {
            writeMeasurementRow(sheet, measurementStyle, multimeterRow, "", "", "", "", "");
        }

        Workbook workbook = sheet.getWorkbook();
        Font sectionFont = workbook.createFont();
        sectionFont.setFontName("Arial");
        sectionFont.setFontHeightInPoints((short) 10);

        CellStyle sectionTextStyle = workbook.createCellStyle();
        sectionTextStyle.setFont(sectionFont);
        sectionTextStyle.setWrapText(true);
        sectionTextStyle.setAlignment(HorizontalAlignment.LEFT);
        sectionTextStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle sectionHeaderStyle = workbook.createCellStyle();
        sectionHeaderStyle.cloneStyleFrom(sectionTextStyle);
        sectionHeaderStyle.setAlignment(HorizontalAlignment.CENTER);
        setThinBorders(sectionHeaderStyle);

        CellStyle sectionCellStyle = workbook.createCellStyle();
        sectionCellStyle.cloneStyleFrom(sectionTextStyle);
        setThinBorders(sectionCellStyle);

        Font sectionSmallFont = workbook.createFont();
        sectionSmallFont.setFontName("Arial");
        sectionSmallFont.setFontHeightInPoints((short) 9);

        CellStyle sectionSmallCenterStyle = workbook.createCellStyle();
        sectionSmallCenterStyle.cloneStyleFrom(sectionCellStyle);
        sectionSmallCenterStyle.setFont(sectionSmallFont);
        sectionSmallCenterStyle.setAlignment(HorizontalAlignment.CENTER);

        int infoRow = multimeterRow + 1;
        setRowHeightPx(sheet, infoRow, 10f);

        int sectionTitleRow = infoRow + 1;
        setRowHeightPx(sheet, sectionTitleRow, 20f);
        setMergedText(sheet, sectionTextStyle, sectionTitleRow, sectionTitleRow, 0, 25,
                "11. Сведения о нормативных документах (НД), регламентирующих значения показателей " +
                        "и НД на методы (методики) измерений:");

        int spacerRow = sectionTitleRow + 1;
        setRowHeightPx(sheet, spacerRow, 10f);

        int sectionHeaderRow = spacerRow + 1;
        setRowHeightPx(sheet, sectionHeaderRow, 40f);
        setMergedText(sheet, sectionHeaderStyle, sectionHeaderRow, sectionHeaderRow, 0, 4,
                "Измеряемый показатель");
        setMergedText(sheet, sectionHeaderStyle, sectionHeaderRow, sectionHeaderRow, 5, 14,
                "Документы, наименование НД, регламентирующих значения характеристик, показателей (к сведению)");
        setMergedText(sheet, sectionHeaderStyle, sectionHeaderRow, sectionHeaderRow, 15, 25,
                "Документы, устанавливающие правила и методы исследований (испытаний) и измерений");

        int sectionRowIndex = sectionHeaderRow + 1;
        int lastSectionRow = sectionHeaderRow;
        if (hasMedSheet) {
            setRowHeightPx(sheet, sectionRowIndex, 96f);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 0, 4,
                    "Мощность дозы гамма-излучения; минимальное значение МЭД гамма-излучения; " +
                            "максимальное значение МЭД гамма-излучения");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 5, 14,
                    "СанПиН 2.6.1.2800-10 \"Гигиенические требования по ограничению облучения населения " +
                            "за счет природных источников ионизирующего излучения\"");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 15, 25,
                    "МР 2.6.1.0333-23 \"Радиационный контроль и санитарно-эпидемиологическая оценка жилых, " +
                            "общественных и производственных зданий и сооружений по показателям радиационной безопасности\" " +
                            "п. IV, V");
            lastSectionRow = sectionRowIndex;
            sectionRowIndex++;
        }

        if (hasEroaSheet) {
            setRowHeightPx(sheet, sectionRowIndex, 150f);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 0, 4,
                    "Эквивалентная равновесная объемная активность (ЭРОА) радона; " +
                            "Эквивалентная равновесная объемная активность (ЭРОА) торона; " +
                            "Среднегодовое значение эквивалентной равновесной объемной активности изотопов радона");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 5, 14,
                    "СанПиН 2.6.1.2800-10 \"Гигиенические требования по ограничению облучения населения " +
                            "за счет природных источников ионизирующего излучения\"");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 15, 25,
                    "МР 2.6.1.0333-23 \"Радиационный контроль и санитарно-эпидемиологическая оценка жилых, " +
                            "общественных и производственных зданий и сооружений по показателям радиационной безопасности\" " +
                            "п. IV, V");
            lastSectionRow = sectionRowIndex;
            sectionRowIndex++;
        }

        boolean hasMicroclimateSheet = hasSheetWithPrefix(sheet.getWorkbook(), "Микроклимат");
        if (hasMicroclimateSheet) {
            setRowHeightPx(sheet, sectionRowIndex, 133f);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 0, 4,
                    "Температура воздуха, скорость движения воздуха, " +
                            "относительная влажность воздуха, результирующая температура");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 5, 14,
                    "СанПиН 1.2.3685-21 \"Гигиенические нормативы и требования к обеспечению безопасности " +
                            "и (или) безвредности для человека факторов среды обитания\"");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 15, 25,
                    "МИ М.08-2021 \"Методика измерений показателей микроклимата\n" +
                            "на рабочих местах в помещениях (сооружениях, кабинах), в\n" +
                            "помещениях жилых зданий (в том числе зданиях общежитий),\n" +
                            "помещениях общественных, административных и бытовых\n" +
                            "зданий (сооружений), помещениях специального подвижного\n" +
                            "состава железнодорожного транспорта и метрополитена, в\n" +
                            "системах вентиляции промышленных, общественных и жилых\n" +
                            "зданий (сооружений), на открытом воздухе\" п.11.2");
            lastSectionRow = sectionRowIndex;
            sectionRowIndex++;
        }

        boolean hasArtificialLightingSheet = sheet.getWorkbook().getSheet("Иск освещение") != null;
        if (hasArtificialLightingSheet) {
            setRowHeightPx(sheet, sectionRowIndex, 84f);
            String lightingIndicatorText = buildLightingSectionIndicatorText(sheet.getWorkbook());
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 0, 4,
                    lightingIndicatorText);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 5, 14,
                    "СанПиН 1.2.3685-21 \"Гигиенические нормативы и требования к обеспечению безопасности " +
                            "и (или) безвредности для человека факторов среды обитания\"");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 15, 25,
                    "МИ СС.09\u22122021 \"Метод измерений показателей световой среды Методика измерений показателей " +
                            "световой среды\nна рабочих местах, в помещениях и оконных конструкциях жилых и общественных " +
                            "зданий (сооружений), селитебной территории\" п. 10.2");
            lastSectionRow = sectionRowIndex;
            sectionRowIndex++;
        }

        if (hasVentilationSheet) {
            VentilationIndicators ventilationIndicators = findVentilationIndicators(sheet.getWorkbook());
            String indicatorText = buildVentilationIndicatorText(ventilationIndicators);
            setRowHeightPx(sheet, sectionRowIndex, 133f);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 0, 4, indicatorText);
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 5, 14,
                    "СП 54.13330.2022 \"СНиП 31-01-2003 Здания жилые многоквартирные\"");
            setMergedText(sheet, sectionSmallCenterStyle, sectionRowIndex, sectionRowIndex, 15, 25,
                    "МИ М.08-2021 \"Методика измерений показателей микроклимата\n" +
                            "на рабочих местах в помещениях (сооружениях, кабинах), в\n" +
                            "помещениях жилых зданий (в том числе зданиях общежитий),\n" +
                            "помещениях общественных, административных и бытовых\n" +
                            "зданий (сооружений), помещениях специального подвижного\n" +
                            "состава железнодорожного транспорта и метрополитена, в\n" +
                            "системах вентиляции промышленных, общественных и жилых\n" +
                            "зданий (сооружений), на открытом воздухе\" п.11.2");
            lastSectionRow = sectionRowIndex;
        }

        int spacerAfterSections = lastSectionRow + 1;
        setRowHeightPx(sheet, spacerAfterSections, 10f);

        int additionalInfoRow = spacerAfterSections + 1;
        setRowHeightPx(sheet, additionalInfoRow, 200f);
        String infoText = (additionalInfoText == null || additionalInfoText.isBlank())
                ? "12. Дополнительные сведения (характеристика объекта):" : additionalInfoText;

        Font infoFont = workbook.createFont();
        infoFont.setFontName("Arial");
        infoFont.setFontHeightInPoints((short) 10);

        Font redFont = workbook.createFont();
        redFont.setFontName("Arial");
        redFont.setFontHeightInPoints((short) 10);
        redFont.setColor(IndexedColors.RED.getIndex());

        CellStyle additionalInfoStyle = workbook.createCellStyle();
        additionalInfoStyle.setFont(infoFont);
        additionalInfoStyle.setWrapText(true);
        additionalInfoStyle.setAlignment(HorizontalAlignment.LEFT);
        additionalInfoStyle.setVerticalAlignment(VerticalAlignment.TOP);

        XSSFRichTextString richInfo = new XSSFRichTextString(infoText);
        richInfo.applyFont(infoFont);
        String highlight = "летний период";
        int highlightIndex = infoText.indexOf(highlight);
        if (highlightIndex >= 0) {
            richInfo.applyFont(highlightIndex, highlightIndex + highlight.length(), redFont);
        }
        setMergedRichText(sheet, additionalInfoStyle, additionalInfoRow, additionalInfoRow, 0, 25, richInfo);

        int spacerAfterAdditionalInfo = additionalInfoRow + 1;
        setRowHeightPx(sheet, spacerAfterAdditionalInfo, 7f);

        int section13Row = spacerAfterAdditionalInfo + 1;
        setRowHeightPx(sheet, section13Row, 20f);
        setMergedText(sheet, sectionTextStyle, section13Row, section13Row, 0, 25,
                "13. Сведения о дополнении, отклонении или исключении из методов: - ");

        int spacerAfterSection13 = section13Row + 1;
        setRowHeightPx(sheet, spacerAfterSection13, 7f);

        int section14Row = spacerAfterSection13 + 1;
        setRowHeightPx(sheet, section14Row, 20f);
        setMergedText(sheet, sectionTextStyle, section14Row, section14Row, 0, 25,
                "14. Эскиз (ситуационный план) места проведения измерений с указанием точек измерений:");

        int legendTitleRow = section14Row + 1;
        setRowHeightPx(sheet, legendTitleRow, 20f);
        setCellValue(sheet, sectionTextStyle, legendTitleRow, 18, "Условные обозначения");

        CellStyle legendBoxStyle = workbook.createCellStyle();
        legendBoxStyle.cloneStyleFrom(sectionTextStyle);
        legendBoxStyle.setAlignment(HorizontalAlignment.CENTER);
        legendBoxStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        legendBoxStyle.setWrapText(false);
        legendBoxStyle.setIndention((short) 0);
        setThinBorders(legendBoxStyle);

        CellStyle legendTextStyle = workbook.createCellStyle();
        legendTextStyle.cloneStyleFrom(sectionTextStyle);
        legendTextStyle.setWrapText(true);
        legendTextStyle.setVerticalAlignment(VerticalAlignment.TOP);

        int legendRow = legendTitleRow + 1;
        setRowHeightPx(sheet, legendRow, 45f);
        setMergedText(sheet, legendBoxStyle, legendRow, legendRow, 18, 20, "поле №1");
        setMergedText(sheet, legendTextStyle, legendRow, legendRow, 21, 25,
                "- Поля для измерения средней горизонтальной освещенности на уровне земли");
    }

    private static void writeHeaderRow(Sheet sheet, int rowIndex, CellStyle headerStyle, String[] headers) {
        setMergedText(sheet, headerStyle, rowIndex, rowIndex, 0, 3, headers[0]);
        setMergedText(sheet, headerStyle, rowIndex, rowIndex, 4, 9, headers[1]);
        setMergedText(sheet, headerStyle, rowIndex, rowIndex, 10, 12, headers[2]);
        setMergedText(sheet, headerStyle, rowIndex, rowIndex, 13, 19, headers[3]);
        setMergedText(sheet, headerStyle, rowIndex, rowIndex, 20, 25, headers[4]);
    }

    private static void writeMeasurementRow(Sheet sheet,
                                            CellStyle style,
                                            int rowIndex,
                                            String indicatorValue,
                                            String instrumentValue,
                                            String errorValue,
                                            String verificationValue) {
        writeMeasurementRow(sheet, style, rowIndex, indicatorValue, instrumentValue, errorValue, verificationValue, "");
    }

    private static void writeMeasurementRow(Sheet sheet,
                                            CellStyle style,
                                            int rowIndex,
                                            String indicatorValue,
                                            String instrumentValue,
                                            String errorValue,
                                            String verificationValue,
                                            String serialValue) {
        setMergedText(sheet, style, rowIndex, rowIndex, 0, 3, indicatorValue);
        setMergedText(sheet, style, rowIndex, rowIndex, 4, 9, instrumentValue);
        setMergedText(sheet, style, rowIndex, rowIndex, 10, 12, serialValue);
        setMergedText(sheet, style, rowIndex, rowIndex, 13, 19, errorValue);
        setMergedText(sheet, style, rowIndex, rowIndex, 20, 25, verificationValue);
    }

    private static String buildLightingIndicatorText(boolean hasLightingIndicator,
                                                     boolean hasKeoIndicator,
                                                     boolean hasPulsation,
                                                     boolean hasStreetLightingIndicator) {
        java.util.List<String> values = new java.util.ArrayList<>();
        if (hasLightingIndicator) {
            values.add("Освещённость");
        }
        if (hasKeoIndicator) {
            values.add("КЕО");
        }
        if (hasPulsation) {
            values.add("коэффициент пульсации");
        }
        if (hasStreetLightingIndicator) {
            values.add("средняя горизонтальная освещенность на уровне земли");
        }
        return String.join(", ", values);
    }

    private static String buildLightingSectionIndicatorText(Workbook workbook) {
        java.util.List<String> values = new java.util.ArrayList<>();
        if (workbook == null) return "";

        boolean hasKeoSheet = workbook.getSheet("КЕО") != null;
        boolean hasStreetLightingSheet = workbook.getSheet("Иск освещение (2)") != null;
        boolean hasPulsation = hasLightingPulsationValues(workbook);

        if (hasKeoSheet) {
            values.add("Освещённость, КЕО");
        } else {
            values.add("Освещённость");
        }
        if (hasStreetLightingSheet) {
            values.add("средняя горизонтальная освещенность на уровне земли");
        }
        if (hasPulsation) {
            values.add("коэффициент пульсации");
        }
        return String.join(", ", values);
    }

    private static String readIndicatorsText(Sheet sheet) {
        if (sheet == null) return "";
        Row row = sheet.getRow(29);
        if (row == null) return "";
        Cell cell = row.getCell(1);
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell);
    }

    private static boolean hasResultantTemperatureValues(Workbook workbook) {
        if (workbook == null) return false;
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) continue;
            String name = sheet.getSheetName();
            if (name == null || !name.startsWith("Микроклимат")) continue;
            int lastRow = sheet.getLastRowNum();
            for (int rowIndex = 5; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;
                Cell cell = row.getCell(10);
                if (cell == null) continue;
                String value = formatter.formatCellValue(cell, evaluator);
                if (value != null && !value.trim().isEmpty()) {
                    return true;
                }
            }
        }
        return false;
    }

    private static boolean hasLightingPulsationValues(Workbook workbook) {
        if (workbook == null) return false;
        Sheet lightingSheet = workbook.getSheet("Иск освещение");
        if (lightingSheet == null) return false;
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        int lastRow = lightingSheet.getLastRowNum();
        for (int rowIndex = 7; rowIndex <= lastRow; rowIndex++) {
            Row row = lightingSheet.getRow(rowIndex);
            if (row == null) continue;
            Cell cell = row.getCell(12);
            if (cell == null) continue;
            String value = formatter.formatCellValue(cell, evaluator);
            if (value != null && !value.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasSheetWithPrefix(Workbook workbook, String prefix) {
        if (workbook == null || prefix == null) return false;
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            String name = workbook.getSheetName(i);
            if (name != null && name.startsWith(prefix)) {
                return true;
            }
        }
        return false;
    }

    private static VentilationIndicators findVentilationIndicators(Workbook workbook) {
        VentilationIndicators indicators = new VentilationIndicators();
        if (workbook == null) return indicators;
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet == null) continue;
            String name = sheet.getSheetName();
            if (name == null || !name.startsWith("Вентиляция")) continue;

            int lastRow = sheet.getLastRowNum();
            for (int rowIndex = 4; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) continue;
                String flowValue = formatter.formatCellValue(row.getCell(9), evaluator);
                if (flowValue == null || flowValue.trim().isEmpty()) continue;

                String placeText = formatter.formatCellValue(row.getCell(2), evaluator);
                if (placeText == null) continue;
                String normalized = placeText.trim().toLowerCase(java.util.Locale.ROOT);
                if (normalized.endsWith("(вытяжка)")) {
                    indicators.hasExhaustRate = true;
                } else if (normalized.endsWith("(приток)")) {
                    indicators.hasSupplyRate = true;
                }
                if (indicators.hasExhaustRate && indicators.hasSupplyRate) {
                    return indicators;
                }
            }
        }

        return indicators;
    }

    private static String buildVentilationIndicatorText(VentilationIndicators indicators) {
        if (indicators == null) {
            return "Скорость воздушного потока, производительность вентсистем, площадь сечения";
        }
        if (indicators.hasExhaustRate && indicators.hasSupplyRate) {
            return "Скорость воздушного потока, производительность вентсистем, " +
                    "кратность воздухообмена по вытяжке, кратность воздухообмена по притоку, площадь сечения";
        }
        if (indicators.hasExhaustRate) {
            return "Скорость воздушного потока, производительность вентсистем, " +
                    "кратность воздухообмена по вытяжке, площадь сечения";
        }
        if (indicators.hasSupplyRate) {
            return "Скорость воздушного потока, производительность вентсистем, " +
                    "кратность воздухообмена по притоку, площадь сечения";
        }
        return "Скорость воздушного потока, производительность вентсистем, площадь сечения";
    }

    private static void setRowHeightPx(Sheet sheet, int rowIndex, float heightPx) {
        if (sheet == null) return;
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        row.setHeightInPoints(heightPx * 72f / 96f);
    }

    private static void setRowHeightPoints(Sheet sheet, int rowIndex, float heightPoints) {
        if (sheet == null) return;
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        row.setHeightInPoints(heightPoints);
    }

    private static float getRowHeightPoints(Sheet sheet, int rowIndex) {
        if (sheet == null) return -1f;
        Row row = sheet.getRow(rowIndex);
        if (row == null) return -1f;
        return row.getHeightInPoints();
    }

    private static void setMergedText(Sheet sheet,
                                      CellStyle style,
                                      int firstRow, int lastRow,
                                      int firstCol, int lastCol,
                                      String text) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);

        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(text);
                }
            }
        }
    }

    private static void setMergedRichText(Sheet sheet,
                                          CellStyle style,
                                          int firstRow, int lastRow,
                                          int firstCol, int lastCol,
                                          XSSFRichTextString text) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);

        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(text);
                }
            }
        }
    }

    private static void setCellValue(Sheet sheet, CellStyle style, int rowIndex, int colIndex, String text) {
        if (sheet == null) return;
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        Cell cell = row.getCell(colIndex);
        if (cell == null) cell = row.createCell(colIndex);
        cell.setCellStyle(style);
        cell.setCellValue(text);
    }

    private static final class VentilationIndicators {
        private boolean hasExhaustRate;
        private boolean hasSupplyRate;
    }
}
