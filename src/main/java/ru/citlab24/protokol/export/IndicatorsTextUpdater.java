package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.*;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

final class IndicatorsTextUpdater {
    private IndicatorsTextUpdater() {}

    static void updateIndicatorsText(Workbook wb) {
        if (wb == null) return;

        Sheet titleSheet = wb.getSheet("Титульная страница");
        if (titleSheet == null) return;

        String indicatorsText = buildIndicatorsText(wb);

        Row row = titleSheet.getRow(29);
        if (row == null) {
            row = titleSheet.createRow(29);
        }
        Cell cell = row.getCell(1);
        if (cell == null) {
            cell = row.createCell(1);
        }
        cell.setCellValue(indicatorsText);
        AllExcelExporter.adjustRowHeightForMergedText(titleSheet, 29, 1, 25, indicatorsText);
    }

    // Стало
    private static String buildIndicatorsText(Workbook wb) {
        String baseText = "Показатели, по которым проводились измерения:";
        if (wb == null) return baseText;

        List<String> indicators = new ArrayList<>();
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();

        boolean hasMicroclimateValues = hasDataInColumnAfterRow(
                wb, "Микроклимат", 5, 5, formatter, evaluator);
        if (hasMicroclimateValues) {
            indicators.add("температура воздуха; относительная влажность воздуха; " +
                    "скорость движения воздуха; результирующая температура;");
        }

        boolean hasVentilationSheet = hasSheetWithPrefix(wb, "Вентиляция");
        if (hasVentilationSheet) {
            indicators.add("скорость воздушного потока; производительность вентсистем; площадь сечения;");
            VentilationIndicators ventilationIndicators = findVentilationIndicators(wb, formatter, evaluator);
            if (ventilationIndicators.hasExhaustRate) {
                indicators.add("кратность воздухообмена по вытяжке;");
            }
            if (ventilationIndicators.hasSupplyRate) {
                indicators.add("кратность воздухообмена по притоку;");
            }
        }

        if (wb.getSheet("МЭД") != null) {
            indicators.add("минимальное значение МЭД гамма-излучения; максимальное значение МЭД гамма-излучения;");
        }

        if (wb.getSheet("МЭД (2)") != null) {
            indicators.add("мощность дозы гамма-излучения;");
        }

        if (wb.getSheet("ЭРОА радона") != null) {
            indicators.add("среднегодовое значение эквивалентной равновесной объемной активности " +
                    "изотопов радона; эквивалентная равновесная объемная активность (ЭРОА) радона; " +
                    "эквивалентная равновесная объемная активность (ЭРОА) торона;");
        }

        Integer lightingIndicatorIndex = null;
        if (wb.getSheet("Иск освещение") != null) {
            lightingIndicatorIndex = indicators.size();
            indicators.add("освещенность (искусственная);");

            // ВАЖНО: пульсацию проверяем ТОЛЬКО на листе "Иск освещение", а не по префиксу
            boolean hasPulsation = hasDataInExactSheetColumnAfterRow(
                    wb, "Иск освещение", 12, 7, formatter, evaluator);
            if (hasPulsation) {
                indicators.add("коэффициент пульсации;");
            }
        }

        if (wb.getSheet("Иск освещение (2)") != null) {
            indicators.add("средняя горизонтальная освещенность на уровне земли;");
        }

        if (wb.getSheet("КЕО") != null) {
            if (lightingIndicatorIndex != null) {
                indicators.set(lightingIndicatorIndex, "освещенность (искусственная, естественная);");
            }
            indicators.add("коэффициент естественной освещенности;");
        }

        if (indicators.isEmpty()) {
            return baseText;
        }

        return baseText + " " + String.join(" ", indicators);
    }

    private static boolean hasSheetWithPrefix(Workbook wb, String prefix) {
        if (wb == null || prefix == null) return false;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetName(i);
            if (name != null && name.startsWith(prefix)) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasDataInColumnAfterRow(Workbook wb,
                                                   String sheetPrefix,
                                                   int columnIndex,
                                                   int startRowIndex,
                                                   DataFormatter formatter,
                                                   FormulaEvaluator evaluator) {
        if (wb == null) return false;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) continue;
            String name = sheet.getSheetName();
            if (name == null || !name.startsWith(sheetPrefix)) continue;
            if (hasDataInColumnAfterRow(sheet, columnIndex, startRowIndex, formatter, evaluator)) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasDataInColumnAfterRow(Sheet sheet,
                                                   int columnIndex,
                                                   int startRowIndex,
                                                   DataFormatter formatter,
                                                   FormulaEvaluator evaluator) {
        if (sheet == null) return false;
        int lastRow = sheet.getLastRowNum();
        for (int rowIndex = startRowIndex; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            Cell cell = row.getCell(columnIndex);
            if (cell == null) continue;
            String value = formatter.formatCellValue(cell, evaluator);
            if (value != null && !value.trim().isEmpty()) {
                return true;
            }
        }
        return false;
    }

    private static VentilationIndicators findVentilationIndicators(Workbook wb,
                                                                   DataFormatter formatter,
                                                                   FormulaEvaluator evaluator) {
        VentilationIndicators indicators = new VentilationIndicators();
        if (wb == null) return indicators;

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
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
                String normalized = placeText.trim().toLowerCase(Locale.ROOT);
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
    //  проверка строго по одному листу (точное имя)
    private static boolean hasDataInExactSheetColumnAfterRow(Workbook wb,
                                                             String sheetName,
                                                             int columnIndex,
                                                             int startRowIndex,
                                                             DataFormatter formatter,
                                                             FormulaEvaluator evaluator) {
        if (wb == null || sheetName == null) return false;
        Sheet sheet = wb.getSheet(sheetName);
        if (sheet == null) return false;
        return hasDataInColumnAfterRow(sheet, columnIndex, startRowIndex, formatter, evaluator);
    }

    private static final class VentilationIndicators {
        private boolean hasExhaustRate;
        private boolean hasSupplyRate;
    }
}
