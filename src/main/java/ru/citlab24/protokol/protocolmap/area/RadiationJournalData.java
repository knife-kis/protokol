package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class RadiationJournalData {
    static final int DEFAULT_POINT_COUNT = 1;

    private RadiationJournalData() {
    }

    static ProtocolData resolveProtocolData(File sourceFile) {
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
        String controlDate = "";

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
            controlDate = resolveControlDate(workbook);
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
                pointCount,
                controlDate
        );
    }

    private static String resolveControlDate(Workbook workbook) {
        if (workbook == null || workbook.getNumberOfSheets() == 0) {
            return "";
        }
        Sheet sheet = workbook.getSheetAt(0);
        if (sheet == null) {
            return "";
        }
        int targetRowIndex = 6;
        Row row = sheet.getRow(targetRowIndex);
        if (row == null) {
            return "";
        }
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        Pattern pattern = Pattern.compile("\\b\\d{1,2}\\s+[а-яА-Я]+\\s+\\d{4}\\s*г?\\.?\\b");
        for (int colIndex = 0; colIndex < row.getLastCellNum(); colIndex++) {
            Cell cell = row.getCell(colIndex);
            String value = formatter.formatCellValue(cell, evaluator).trim();
            if (value.isEmpty()) {
                continue;
            }
            Matcher matcher = pattern.matcher(value);
            if (matcher.find()) {
                return matcher.group().trim();
            }
        }
        return "";
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
        String sectionHeader = "Сведения о нормативных документах (НД), " +
                "регламентирующих значения показателей и НД на методы (методики) измерений:";
        DataFormatter formatter = new DataFormatter();
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        boolean inSection = false;
        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            if (!inSection) {
                for (Cell cell : row) {
                    String value = formatter.formatCellValue(cell, evaluator).trim();
                    if (value.contains(sectionHeader)) {
                        inSection = true;
                        break;
                    }
                }
                if (!inSection) {
                    continue;
                }
            }
            for (Cell cell : row) {
                String value = formatter.formatCellValue(cell, evaluator).trim();
                if (!value.contains("Мощность дозы гамма-излучения")) {
                    continue;
                }
                int baseCol = cell.getColumnIndex();
                CellRangeAddress currentRegion = findMergedRegion(sheet, row.getRowNum(), baseCol);
                int nextCol = currentRegion != null ? currentRegion.getLastColumn() + 1 : baseCol + 1;
                CellRangeAddress nextRegion = findMergedRegion(sheet, row.getRowNum(), nextCol);
                int nextNextCol = nextRegion != null ? nextRegion.getLastColumn() + 1 : nextCol + 1;
                String rightValue = readMergedCellValue(sheet, row.getRowNum(), nextCol, formatter, evaluator);
                String rightRightValue = readMergedCellValue(sheet, row.getRowNum(), nextNextCol, formatter, evaluator);
                if (!rightRightValue.isBlank()) {
                    return rightRightValue;
                }
                if (!rightValue.isBlank()) {
                    return rightValue;
                }
                return "";
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

    record ProtocolData(String preparedLine,
                        String observationPeriod,
                        String settlement,
                        String method,
                        String customerName,
                        String customerRequest,
                        String customerPresence,
                        String observationLocation,
                        int pointCount,
                        String controlDate) {
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
                    DEFAULT_POINT_COUNT,
                    ""
            );
        }
    }
}
