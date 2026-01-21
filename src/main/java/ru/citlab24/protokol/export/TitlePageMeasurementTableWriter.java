package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

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

    static void rebuildTable(Workbook wb) {
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
        write(sheet, headerRow, headerStyle, measurementStyle);

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


    static void write(Sheet sheet, int headerRow, CellStyle headerStyle, CellStyle measurementStyle) {
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
                "С-АШ/22-01-2025/403727315  до 21.01.2026");

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
        setRowHeightPx(sheet, eroaRow, 136f);
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
}
