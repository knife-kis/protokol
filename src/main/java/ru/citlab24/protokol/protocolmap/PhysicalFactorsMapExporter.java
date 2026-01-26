package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Locale;

public final class PhysicalFactorsMapExporter {
    private static final String REGISTRATION_PREFIX = "Регистрационный номер карты замеров:";
    private static final double COLUMN_WIDTH_SCALE = 0.9;
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;
    private static final double A4_LANDSCAPE_HEIGHT_CM = 21.0;
    private static final int MICROCLIMATE_SOURCE_START_ROW = 5;
    private static final int MICROCLIMATE_SOURCE_MERGED_LAST_COL = 21;
    private static final int MICROCLIMATE_BLOCK_SIZE = 3;
    private static final int[] MICROCLIMATE_TOP_BOTTOM_COLUMNS = {3, 5, 7, 9, 10, 12};
    private static final int MICROCLIMATE_AIR_SPEED_START_COL = 13;
    private static final int MICROCLIMATE_AIR_SPEED_END_COL = 15;
    private static final int MED2_SOURCE_START_ROW = 5;
    private static final int VENTILATION_SOURCE_START_ROW = 4;
    private static final int VENTILATION_TARGET_START_ROW = 3;
    private static final int VENTILATION_LAST_COL = 9;

    private PhysicalFactorsMapExporter() {
    }

    public static File generateMap(File sourceFile) throws IOException {
        String registrationNumber = resolveRegistrationNumber(sourceFile);
        MapHeaderData headerData = resolveHeaderData(sourceFile);
        java.util.List<String> measurementDates = extractMeasurementDatesList(headerData.measurementDates);
        String protocolNumber = resolveProtocolNumber(sourceFile);
        String contractText = resolveContractText(sourceFile);
        String measurementPerformer = resolveMeasurementPerformer(sourceFile);
        String controlDate = resolveControlDate(sourceFile);
        String specialConditions = resolveSpecialConditions(sourceFile);
        String measurementMethods = resolveMeasurementMethods(sourceFile);
        java.util.List<InstrumentData> instruments = resolveMeasurementInstruments(sourceFile);
        boolean hasMicroclimateSheet = hasSheetWithPrefix(sourceFile, "Микроклимат");
        boolean hasMedSheet = hasSheetWithName(sourceFile, "МЭД");
        boolean hasMed2Sheet = hasSheetWithName(sourceFile, "МЭД (2)");
        File targetFile = buildTargetFile(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("карта замеров");
            applySheetDefaults(workbook, sheet);
            applyHeaders(sheet, registrationNumber);
            createTitleRows(workbook, sheet, registrationNumber, headerData, measurementPerformer, controlDate);
            createSecondPageRows(workbook, sheet, protocolNumber, contractText, headerData,
                    specialConditions, measurementMethods, instruments);
            int microclimateDataStartRow = PhysicalFactorsMapResultsTabBuilder.createResultsSheet(
                    workbook, measurementDates, hasMicroclimateSheet);
            Sheet resultsSheet = workbook.getSheet("Микроклимат");
            int medDataStartRow = -1;
            Sheet medSheet = null;
            int med2DataStartRow = -1;
            Sheet med2Sheet = null;
            Sheet ventilationSheet = VentilationMapTabBuilder.createSheet(workbook);
            if (hasMedSheet) {
                medDataStartRow = MedMapTabBuilder.createMedResultsSheet(workbook);
                medSheet = workbook.getSheet("МЭД");
            }
            if (hasMed2Sheet) {
                med2DataStartRow = Med2MapTabBuilder.createMed2ResultsSheet(workbook);
                med2Sheet = workbook.getSheet("МЭД (2)");
            }
            if (resultsSheet != null) {
                applyHeaders(resultsSheet, registrationNumber);
            }
            if (medSheet != null) {
                applyHeaders(medSheet, registrationNumber);
            }
            if (med2Sheet != null) {
                applyHeaders(med2Sheet, registrationNumber);
            }
            if (ventilationSheet != null) {
                applyHeaders(ventilationSheet, registrationNumber);
            }
            if (hasMicroclimateSheet) {
                fillMicroclimateResults(sourceFile, workbook, resultsSheet, microclimateDataStartRow);
            }
            if (hasMedSheet) {
                fillMedResults(sourceFile, workbook, medSheet, medDataStartRow);
            }
            if (hasMed2Sheet) {
                fillMed2Results(sourceFile, workbook, med2Sheet, med2DataStartRow);
            }
            fillVentilationResults(sourceFile, workbook, ventilationSheet);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                workbook.write(out);
            }
        }

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

    private static boolean hasSheetWithPrefix(File sourceFile, String prefix) {
        if (sourceFile == null || !sourceFile.exists()) {
            return false;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            int count = workbook.getNumberOfSheets();
            for (int i = 0; i < count; i++) {
                String name = workbook.getSheetName(i);
                if (name != null && name.startsWith(prefix)) {
                    return true;
                }
            }
        } catch (Exception ex) {
            return false;
        }
        return false;
    }

    private static boolean hasSheetWithName(File sourceFile, String name) {
        if (sourceFile == null || !sourceFile.exists()) {
            return false;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            return workbook.getSheet(name) != null;
        } catch (Exception ex) {
            return false;
        }
    }

    private static File buildTargetFile(File sourceFile) {
        String name = sourceFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = dotIndex > 0 ? name.substring(0, dotIndex) : name;
        return new File(sourceFile.getParentFile(), baseName + "_карта.xlsx");
    }

    private static void fillMicroclimateResults(File sourceFile,
                                                Workbook targetWorkbook,
                                                Sheet targetSheet,
                                                int targetStartRow) {
        if (sourceFile == null || !sourceFile.exists() || targetSheet == null || targetStartRow < 0) {
            return;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {
            Sheet sourceSheet = findSheetWithPrefix(sourceWorkbook, "Микроклимат");
            if (sourceSheet == null) {
                return;
            }
            DataFormatter formatter = new DataFormatter();
            CellStyle centerStyle = createMicroclimateDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            CellStyle leftStyle = createMicroclimateDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
            CellStyle mergedRowStyle = createMicroclimateDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);

            int sourceRowIndex = MICROCLIMATE_SOURCE_START_ROW;
            int targetRowIndex = targetStartRow;
            int targetDataStartRow = targetStartRow;
            int lastRow = sourceSheet.getLastRowNum();
            PageBreakHelper pageBreakHelper = new PageBreakHelper(targetSheet, targetDataStartRow);

            while (sourceRowIndex <= lastRow) {
                CellRangeAddress mergedRow = findMergedRegion(sourceSheet, sourceRowIndex, 0);
                if (isMicroclimateMergedRow(mergedRow, sourceRowIndex)) {
                    String text = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter);
                    if (text.isBlank() && !hasRowContent(sourceSheet, sourceRowIndex, formatter)) {
                        break;
                    }
                    pageBreakHelper.ensureSpace(targetRowIndex, 1);
                    mergeCellRangeWithValue(targetSheet, targetRowIndex, targetRowIndex, 0, 15, text, mergedRowStyle);
                    pageBreakHelper.consume(1);
                    targetRowIndex++;
                    sourceRowIndex++;
                    continue;
                }

                if (!hasMicroclimateBlockContent(sourceSheet, sourceRowIndex, formatter)) {
                    break;
                }
                pageBreakHelper.ensureSpace(targetRowIndex, MICROCLIMATE_BLOCK_SIZE);

                String blockLabel = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter);
                String blockPlace = readMergedCellValue(sourceSheet, sourceRowIndex, 1, formatter);
                int targetBlockStart = targetRowIndex;
                int targetBlockEnd = targetRowIndex + MICROCLIMATE_BLOCK_SIZE - 1;

                mergeCellRangeWithValue(targetSheet, targetBlockStart, targetBlockEnd, 0, 0, blockLabel, leftStyle);
                mergeCellRangeWithValue(targetSheet, targetBlockStart, targetBlockEnd, 1, 1, blockPlace, leftStyle);
                mergeCellRangeWithValue(targetSheet, targetBlockStart, targetBlockEnd, 6, 6, "-", centerStyle);

                for (int offset = 0; offset < MICROCLIMATE_BLOCK_SIZE; offset++) {
                    int sourceRow = sourceRowIndex + offset;
                    int targetRow = targetRowIndex + offset;
                    String heightValue = readCellValue(sourceSheet, sourceRow, 4, formatter);
                    String temperatureValue = readCellValue(sourceSheet, sourceRow, 6, formatter);

                    setCellValue(targetSheet, targetRow, 2, heightValue, centerStyle);
                    setCellValue(targetSheet, targetRow, 4, temperatureValue, centerStyle);
                    setCellValue(targetSheet, targetRow, 8, "±", centerStyle);
                    setCellValue(targetSheet, targetRow, 11, "±", centerStyle);
                    mergeAirSpeedCells(targetSheet, targetRow, centerStyle);
                }

                pageBreakHelper.consume(MICROCLIMATE_BLOCK_SIZE);
                targetRowIndex += MICROCLIMATE_BLOCK_SIZE;
                sourceRowIndex += MICROCLIMATE_BLOCK_SIZE;
            }

            if (targetRowIndex > targetDataStartRow) {
                applyPlusAdjacentBorders(targetSheet, targetWorkbook, targetDataStartRow, targetRowIndex - 1);
                applyMicroclimateDataBorders(targetSheet, targetWorkbook, targetDataStartRow, targetRowIndex - 1);
            }
        } catch (Exception ex) {
            // ignore
        }
    }

    private static boolean isMicroclimateMergedRow(CellRangeAddress region, int rowIndex) {
        return region != null
                && region.getFirstRow() == rowIndex
                && region.getLastRow() == rowIndex
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= MICROCLIMATE_SOURCE_MERGED_LAST_COL;
    }

    private static void fillVentilationResults(File sourceFile,
                                               Workbook targetWorkbook,
                                               Sheet targetSheet) {
        if (sourceFile == null || !sourceFile.exists() || targetSheet == null) {
            return;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {

            Sheet sourceSheet = findSheetWithKeyword(sourceWorkbook, "вентиляция");
            if (sourceSheet == null) {
                return;
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();

            CellStyle centerStyle = createVentilationDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            CellStyle leftStyle = createVentilationDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);
            CellStyle mergedRowStyle = createVentilationDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

            java.util.Map<BorderKey, CellStyle> styleCache = new java.util.HashMap<>();

            // ВАЖНО: у тебя VENTILATION_SOURCE_START_ROW = 4 => это POI-индекс строки (0-based),
            // то есть Excel-строка 5. (Excel-строка 4 — оглавление — пропускаем)
            int sourceRowIndex = VENTILATION_SOURCE_START_ROW;
            int targetRowIndex = VENTILATION_TARGET_START_ROW;
            int lastRow = sourceSheet.getLastRowNum();

            boolean started = false;
            int emptyAStreak = 0;

            while (sourceRowIndex <= lastRow) {

                String aValue = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);

                // условие завершения: после того как начали, если подряд много пустых A — таблица закончилась
                if (normalizeText(aValue).isBlank()) {
                    emptyAStreak++;
                    if (started && emptyAStreak >= 20) {
                        break;
                    }
                } else {
                    emptyAStreak = 0;
                }

                // 1) merged-строка (этаж) — переносим целиком
                CellRangeAddress mergedRow = findMergedRegion(sourceSheet, sourceRowIndex, 0);
                if (isVentilationMergedRow(mergedRow, sourceRowIndex)) {
                    String text = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);
                    if (!text.isBlank()) {
                        mergeCellRangeWithValue(targetSheet, targetRowIndex, targetRowIndex,
                                0, VENTILATION_LAST_COL, text, mergedRowStyle);
                        targetRowIndex++;
                        started = true;
                    }
                    sourceRowIndex++;
                    continue;
                }

                // 2) обычная строка: число в A
                if (!startsWithDigit(aValue)) {
                    sourceRowIndex++;
                    continue;
                }

                // C (карта) = C (протокол)
                String placeValue = readMergedCellValue(sourceSheet, sourceRowIndex, 2, formatter, evaluator);

                // I (карта) = I (протокол), если пусто => "-"
                String volumeValue = readMergedCellValue(sourceSheet, sourceRowIndex, 8, formatter, evaluator);
                String normalizedVolume = normalizeText(volumeValue);
                if (normalizedVolume.isBlank() || "-".equals(normalizedVolume)) {
                    volumeValue = "-";
                }

                // J (карта): если I="-", то J="-", иначе J пусто
                String exchangeValue = "-".equals(volumeValue) ? "-" : "";

                // A..J по твоей логике
                setCellValue(targetSheet, targetRowIndex, 0, aValue, centerStyle);         // A
                setCellValue(targetSheet, targetRowIndex, 1, "-", centerStyle);           // B
                setCellValue(targetSheet, targetRowIndex, 2, placeValue, leftStyle);      // C
                setCellValue(targetSheet, targetRowIndex, 3, "", centerStyle);            // D
                setCellValue(targetSheet, targetRowIndex, 4, "±", centerStyle);           // E
                setCellValue(targetSheet, targetRowIndex, 5, "", centerStyle);            // F
                setCellValue(targetSheet, targetRowIndex, 6, "", centerStyle);            // G
                setCellValue(targetSheet, targetRowIndex, 7, "", centerStyle);            // H
                setCellValue(targetSheet, targetRowIndex, 8, volumeValue, centerStyle);   // I
                setCellValue(targetSheet, targetRowIndex, 9, exchangeValue, centerStyle); // J

                // Грани:
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 0, true, true, true, true); // A
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 1, true, true, true, true); // B
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 2, true, true, true, true); // C

                // D: слева/сверху/снизу, справа НЕТ
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 3, true, true, true, false);

                // E: только сверху/снизу
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 4, true, true, false, false);

                // F: сверху/снизу/справа, слева НЕТ
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 5, true, true, false, true);

                // G/H/I/J: со всех сторон
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 6, true, true, true, true);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 7, true, true, true, true);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 8, true, true, true, true);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 9, true, true, true, true);

                targetRowIndex++;
                sourceRowIndex++;
                started = true;
            }

        } catch (Exception ex) {
            // ignore
        }
    }

    private static void fillMedResults(File sourceFile,
                                       Workbook targetWorkbook,
                                       Sheet targetSheet,
                                       int targetStartRow) {
        if (sourceFile == null || !sourceFile.exists() || targetSheet == null || targetStartRow < 0) {
            return;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {
            Sheet sourceSheet = findSheetWithName(sourceWorkbook, "МЭД");
            if (sourceSheet == null) {
                return;
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();

            CellStyle centerStyle = createMedDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            CellStyle leftStyle = createMedDataStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

            int sourceRowIndex = 7;
            int targetRowIndex = targetStartRow;
            int lastRow = sourceSheet.getLastRowNum();
            int emptyAStreak = 0;
            boolean started = false;
            float dataRowHeight = targetSheet.getDefaultRowHeightInPoints() * 3f;

            while (sourceRowIndex <= lastRow) {
                String aValue = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);
                String bValue = readMergedCellValue(sourceSheet, sourceRowIndex, 1, formatter, evaluator);

                if (normalizeText(aValue).isBlank()) {
                    emptyAStreak++;
                    if (started && emptyAStreak >= 10) {
                        break;
                    }
                    sourceRowIndex++;
                    continue;
                }
                emptyAStreak = 0;
                started = true;

                ensureRow(targetSheet, targetRowIndex).setHeightInPoints(dataRowHeight);
                setCellValue(targetSheet, targetRowIndex, 0, aValue, centerStyle);
                setCellValue(targetSheet, targetRowIndex, 1, bValue, leftStyle);
                setCellValue(targetSheet, targetRowIndex, 2, "", centerStyle);

                targetRowIndex++;
                sourceRowIndex++;
            }
        } catch (Exception ex) {
            // ignore
        }
    }

    private static void fillMed2Results(File sourceFile,
                                        Workbook targetWorkbook,
                                        Sheet targetSheet,
                                        int targetStartRow) {
        if (sourceFile == null || !sourceFile.exists() || targetSheet == null || targetStartRow < 0) {
            return;
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {
            Sheet sourceSheet = findSheetWithName(sourceWorkbook, "МЭД (2)");
            if (sourceSheet == null) {
                return;
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();

            CellStyle centerStyle = createMed2BaseStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
            CellStyle leftStyle = createMed2BaseStyle(targetWorkbook,
                    org.apache.poi.ss.usermodel.HorizontalAlignment.LEFT);

            java.util.Map<BorderKey, CellStyle> styleCache = new java.util.HashMap<>();

            int sourceRowIndex = MED2_SOURCE_START_ROW;
            int targetRowIndex = targetStartRow;
            int lastRow = sourceSheet.getLastRowNum();
            int emptyAStreak = 0;
            boolean started = false;

            while (sourceRowIndex <= lastRow) {
                CellRangeAddress mergedRow = findMergedRegion(sourceSheet, sourceRowIndex, 0);
                if (isMed2MergedRow(mergedRow, sourceRowIndex)) {
                    String text = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);
                    if (text.isBlank() && !hasRowContent(sourceSheet, sourceRowIndex, formatter)) {
                        if (started) {
                            break;
                        }
                        sourceRowIndex++;
                        continue;
                    }
                    mergeCellRangeWithValue(targetSheet, targetRowIndex, targetRowIndex, 0, 4, text, centerStyle);
                    targetRowIndex++;
                    sourceRowIndex++;
                    started = true;
                    emptyAStreak = 0;
                    continue;
                }

                String aValue = readMergedCellValue(sourceSheet, sourceRowIndex, 0, formatter, evaluator);
                if (normalizeText(aValue).isBlank()) {
                    emptyAStreak++;
                    if (started && emptyAStreak >= 10) {
                        break;
                    }
                    sourceRowIndex++;
                    continue;
                }
                emptyAStreak = 0;
                started = true;

                String bValue = readMergedCellValue(sourceSheet, sourceRowIndex, 1, formatter, evaluator);

                setCellValue(targetSheet, targetRowIndex, 0, aValue, centerStyle);
                setCellValue(targetSheet, targetRowIndex, 1, bValue, leftStyle);
                setCellValue(targetSheet, targetRowIndex, 2, "", centerStyle);
                setCellValue(targetSheet, targetRowIndex, 3, "±", centerStyle);
                setCellValue(targetSheet, targetRowIndex, 4, "", centerStyle);

                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 0, true, true, true, true);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 1, true, true, true, true);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 2, true, true, true, false);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 3, true, true, false, false);
                applyBorderToCell(targetSheet, targetWorkbook, styleCache, targetRowIndex, 4, true, true, false, true);

                targetRowIndex++;
                sourceRowIndex++;
            }
        } catch (Exception ex) {
            // ignore
        }
    }


    private static boolean isVentilationMergedRow(CellRangeAddress region, int rowIndex) {
        return region != null
                && region.getFirstRow() == rowIndex
                && region.getLastRow() == rowIndex
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= VENTILATION_LAST_COL;
    }

    private static boolean isMed2MergedRow(CellRangeAddress region, int rowIndex) {
        return region != null
                && region.getFirstRow() == rowIndex
                && region.getLastRow() == rowIndex
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= 4;
    }

    private static boolean startsWithDigit(String value) {
        if (value == null) {
            return false;
        }
        String trimmed = value.trim();
        return !trimmed.isEmpty() && Character.isDigit(trimmed.charAt(0));
    }

    private static boolean hasMicroclimateBlockContent(Sheet sheet, int startRow, DataFormatter formatter) {
        for (int offset = 0; offset < MICROCLIMATE_BLOCK_SIZE; offset++) {
            int rowIndex = startRow + offset;
            if (!readMergedCellValue(sheet, rowIndex, 0, formatter).isBlank()) {
                return true;
            }
            if (!readMergedCellValue(sheet, rowIndex, 1, formatter).isBlank()) {
                return true;
            }
            if (!readCellValue(sheet, rowIndex, 4, formatter).isBlank()) {
                return true;
            }
            if (!readCellValue(sheet, rowIndex, 6, formatter).isBlank()) {
                return true;
            }
        }
        return false;
    }

    private static boolean hasRowContent(Sheet sheet, int rowIndex, DataFormatter formatter) {
        if (sheet == null) {
            return false;
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return false;
        }
        for (Cell cell : row) {
            if (!normalizeText(formatter.formatCellValue(cell)).isBlank()) {
                return true;
            }
        }
        return false;
    }


    private static String readCellValue(Sheet sheet, int rowIndex, int colIndex, DataFormatter formatter) {
        if (sheet == null) {
            return "";
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
    }

    private static CellStyle createMicroclimateDataStyle(Workbook workbook,
                                                         org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        return style;
    }

    private static CellStyle createMedDataStyle(Workbook workbook,
                                                org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        return style;
    }

    private static CellStyle createVentilationDataStyle(Workbook workbook,
                                                        org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle createMed2BaseStyle(Workbook workbook,
                                                 org.apache.poi.ss.usermodel.HorizontalAlignment alignment) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(alignment);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle createMicroclimateAdjacentBorderStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setWrapText(true);
        style.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        style.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);
        style.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        style.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
        return style;
    }

    private static void setCellValue(Sheet sheet, int rowIndex, int colIndex, String value, CellStyle style) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
        cell.setCellStyle(style);
    }

    private static Row ensureRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    private static void mergeCellRangeWithValue(Sheet sheet,
                                                int rowStart,
                                                int rowEnd,
                                                int colStart,
                                                int colEnd,
                                                String value,
                                                CellStyle style) {
        for (int r = rowStart; r <= rowEnd; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                row = sheet.createRow(r);
            }
            for (int c = colStart; c <= colEnd; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) {
                    cell = row.createCell(c);
                }
                cell.setCellStyle(style);
            }
        }
        setCellValue(sheet, rowStart, colStart, value, style);
        CellRangeAddress region = new CellRangeAddress(rowStart, rowEnd, colStart, colEnd);
        sheet.addMergedRegion(region);
        applyThinBorders(sheet, region);
    }

    private static void applyThinBorders(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN, region, sheet);
    }

    private static void applyPlusAdjacentBorders(Sheet sheet,
                                                 Workbook workbook,
                                                 int startRow,
                                                 int endRow) {
        DataFormatter formatter = new DataFormatter();
        CellStyle adjacentStyle = createMicroclimateAdjacentBorderStyle(workbook);
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            if (isPlusCell(sheet, rowIndex, 8, formatter)) {
                applyTopBottomBorder(sheet, rowIndex, 7, adjacentStyle);
                applyTopBottomBorder(sheet, rowIndex, 9, adjacentStyle);
            }
            if (isPlusCell(sheet, rowIndex, 11, formatter)) {
                applyTopBottomBorder(sheet, rowIndex, 10, adjacentStyle);
                applyTopBottomBorder(sheet, rowIndex, 12, adjacentStyle);
            }
        }
    }

    private static boolean isPlusCell(Sheet sheet, int rowIndex, int colIndex, DataFormatter formatter) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return false;
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            return false;
        }
        String value = normalizeText(formatter.formatCellValue(cell));
        return "±".equals(value);
    }

    private static void applyTopBottomBorder(Sheet sheet, int rowIndex, int colIndex, CellStyle borderStyle) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        CellStyle currentStyle = cell.getCellStyle();
        if (currentStyle == null || currentStyle.getIndex() == 0) {
            cell.setCellStyle(borderStyle);
        }
    }

    private static void applyMicroclimateDataBorders(Sheet sheet,
                                                     Workbook workbook,
                                                     int startRow,
                                                     int endRow) {
        java.util.Map<BorderKey, CellStyle> styleCache = new java.util.HashMap<>();
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            for (int colIndex : MICROCLIMATE_TOP_BOTTOM_COLUMNS) {
                boolean needsLeft = colIndex == 10;
                boolean needsRight = false;
                applyBorderToCell(sheet, workbook, styleCache, rowIndex, colIndex,
                        true, true, needsLeft, needsRight);
            }
        }
    }

    private static void mergeAirSpeedCells(Sheet sheet, int rowIndex, CellStyle style) {
        CellRangeAddress existing = findMergedRegion(sheet, rowIndex, MICROCLIMATE_AIR_SPEED_START_COL);
        if (existing != null
                && existing.getFirstRow() == rowIndex
                && existing.getLastRow() == rowIndex
                && existing.getFirstColumn() <= MICROCLIMATE_AIR_SPEED_START_COL
                && existing.getLastColumn() >= MICROCLIMATE_AIR_SPEED_END_COL) {
            return;
        }
        mergeCellRangeWithValue(sheet, rowIndex, rowIndex,
                MICROCLIMATE_AIR_SPEED_START_COL, MICROCLIMATE_AIR_SPEED_END_COL, "", style);
    }

    private static void applyBorderToCell(Sheet sheet,
                                          Workbook workbook,
                                          java.util.Map<BorderKey, CellStyle> styleCache,
                                          int rowIndex,
                                          int colIndex,
                                          boolean top,
                                          boolean bottom,
                                          boolean left,
                                          boolean right) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        CellStyle baseStyle = cell.getCellStyle();
        short baseIndex = baseStyle == null ? 0 : baseStyle.getIndex();
        BorderKey key = new BorderKey(baseIndex, top, bottom, left, right);
        CellStyle mergedStyle = styleCache.get(key);
        if (mergedStyle == null) {
            mergedStyle = workbook.createCellStyle();
            if (baseStyle != null && baseStyle.getIndex() != 0) {
                mergedStyle.cloneStyleFrom(baseStyle);
            }
            if (top) {
                mergedStyle.setBorderTop(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            }
            if (bottom) {
                mergedStyle.setBorderBottom(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            }
            if (left) {
                mergedStyle.setBorderLeft(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            }
            if (right) {
                mergedStyle.setBorderRight(org.apache.poi.ss.usermodel.BorderStyle.THIN);
            }
            styleCache.put(key, mergedStyle);
        }
        cell.setCellStyle(mergedStyle);
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

    private static final class BorderKey {
        private final short baseStyleIndex;
        private final boolean top;
        private final boolean bottom;
        private final boolean left;
        private final boolean right;

        private BorderKey(short baseStyleIndex, boolean top, boolean bottom, boolean left, boolean right) {
            this.baseStyleIndex = baseStyleIndex;
            this.top = top;
            this.bottom = bottom;
            this.left = left;
            this.right = right;
        }

        @Override
        public boolean equals(Object obj) {
            if (this == obj) {
                return true;
            }
            if (obj == null || getClass() != obj.getClass()) {
                return false;
            }
            BorderKey other = (BorderKey) obj;
            return baseStyleIndex == other.baseStyleIndex
                    && top == other.top
                    && bottom == other.bottom
                    && left == other.left
                    && right == other.right;
        }

        @Override
        public int hashCode() {
            int result = baseStyleIndex;
            result = 31 * result + (top ? 1 : 0);
            result = 31 * result + (bottom ? 1 : 0);
            result = 31 * result + (left ? 1 : 0);
            result = 31 * result + (right ? 1 : 0);
            return result;
        }
    }

    private static final class PageBreakHelper {
        private static final double POINTS_PER_INCH = 72.0;
        private static final double A4_HEIGHT_POINTS = 842.0;
        private static final double A4_WIDTH_POINTS = 595.0;

        private final Sheet sheet;
        private final int dataStartRow;
        private final double rowHeightPoints;
        private final int firstPageRowsPerPage;
        private final int nextPageRowsPerPage;
        private int rowsOnPage = 0;
        private boolean firstPage = true;

        private PageBreakHelper(Sheet sheet, int dataStartRow) {
            this.sheet = sheet;
            this.dataStartRow = dataStartRow;
            double defaultHeight = sheet.getDefaultRowHeightInPoints();
            this.rowHeightPoints = defaultHeight > 0 ? defaultHeight : 15.0;
            this.firstPageRowsPerPage = calculateRowsPerPage(firstPageHeaderHeight());
            this.nextPageRowsPerPage = calculateRowsPerPage(repeatingHeaderHeight());
        }

        void ensureSpace(int rowIndex, int neededRows) {
            int rowsPerPage = rowsPerPage();
            if (rowIndex < dataStartRow || rowsPerPage <= 0 || neededRows > rowsPerPage) {
                return;
            }
            if (rowsOnPage + neededRows > rowsPerPage) {
                sheet.setRowBreak(rowIndex - 1);
                rowsOnPage = 0;
                firstPage = false;
            }
        }

        void consume(int rows) {
            if (rowsPerPage() <= 0) {
                return;
            }
            rowsOnPage += rows;
        }

        private int rowsPerPage() {
            int rowsPerPage = firstPage ? firstPageRowsPerPage : nextPageRowsPerPage;
            if (rowsPerPage <= 0 && firstPageRowsPerPage > 0) {
                return firstPageRowsPerPage;
            }
            return rowsPerPage;
        }

        private int calculateRowsPerPage(double headerHeight) {
            double marginPoints = (sheet.getMargin(Sheet.TopMargin) + sheet.getMargin(Sheet.BottomMargin)) * POINTS_PER_INCH;
            double pageHeight = pageHeightPoints();
            double available = pageHeight - marginPoints - headerHeight;
            if (available <= 0) {
                return 0;
            }
            return (int) Math.floor(available / rowHeightPoints);
        }

        private double firstPageHeaderHeight() {
            double headerHeight = 0.0;
            for (int i = 0; i < dataStartRow; i++) {
                headerHeight += ensureRow(sheet, i).getHeightInPoints();
            }
            return headerHeight;
        }

        private double repeatingHeaderHeight() {
            CellRangeAddress repeating = sheet.getRepeatingRows();
            if (repeating == null) {
                return 0.0;
            }
            int start = Math.max(0, repeating.getFirstRow());
            int end = Math.max(start, repeating.getLastRow());
            double headerHeight = 0.0;
            for (int i = start; i <= end; i++) {
                headerHeight += ensureRow(sheet, i).getHeightInPoints();
            }
            return headerHeight;
        }

        private double pageHeightPoints() {
            PrintSetup ps = sheet.getPrintSetup();
            if (ps != null && ps.getPaperSize() == PrintSetup.A4_PAPERSIZE) {
                return ps.getLandscape() ? A4_WIDTH_POINTS : A4_HEIGHT_POINTS;
            }
            return A4_WIDTH_POINTS;
        }
    }

    private static Sheet findSheetWithPrefix(Workbook workbook, String prefix) {
        if (workbook == null) {
            return null;
        }
        int count = workbook.getNumberOfSheets();
        for (int i = 0; i < count; i++) {
            String name = workbook.getSheetName(i);
            if (name != null && name.startsWith(prefix)) {
                return workbook.getSheetAt(i);
            }
        }
        return null;
    }

    private static Sheet findSheetWithKeyword(Workbook workbook, String keyword) {
        if (workbook == null || keyword == null) {
            return null;
        }
        String needle = keyword.toLowerCase(Locale.ROOT);
        int count = workbook.getNumberOfSheets();
        for (int i = 0; i < count; i++) {
            String name = workbook.getSheetName(i);
            if (name == null) {
                continue;
            }
            String normalized = normalizeText(name).toLowerCase(Locale.ROOT);
            if (normalized.contains(needle)) {
                return workbook.getSheetAt(i);
            }
        }
        return null;
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
        String datesValue = safe(headerData.measurementDates);
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

        setMergedCellValue(sheet, rowIndex, "5. Дополнительные сведения", plainStyle);
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

        setMergedCellValue(sheet, rowIndex, "6. Эскиз", plainStyle);
        rowIndex++;
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

        // Проставляем стиль во всех ячейках диапазона, а значение — в первой
        for (int col = firstCol; col <= lastCol; col++) {
            Cell c = row.getCell(col);
            if (c == null) c = row.createCell(col);
            c.setCellStyle(style);
            if (col == firstCol) {
                c.setCellValue(value);
            }
        }

        // Для merged-регионов границы надёжнее “прожимать” через RegionUtil
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
        int from = start + MEASUREMENT_DATES_PHRASE.length();
        String tail = text.substring(from).trim();
        if (tail.isEmpty()) {
            return "";
        }
        java.util.LinkedHashSet<String> dates = new java.util.LinkedHashSet<>();
        java.util.regex.Matcher matcher = DATE_PATTERN.matcher(tail);
        while (matcher.find()) {
            dates.add(matcher.group());
        }
        if (dates.isEmpty()) {
            return "";
        }
        return String.join(", ", dates);
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

    private static String extractRepresentative(String text) {
        if (text == null) {
            return "";
        }
        int index = text.indexOf(REPRESENTATIVE_PREFIX);
        if (index < 0) {
            return "";
        }
        return text.substring(index + REPRESENTATIVE_PREFIX.length()).trim();
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
            java.util.regex.Matcher matcher = CONTROL_DATE_PATTERN.matcher(text);
            if (matcher.find()) {
                return matcher.group();
            }
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
                conditions.add("заказчик сообщил, что перед измерением ЭРОА радона " +
                        "здание выдерженно в течении более 12 часов при закрытых дверях и окнах");
            }
            if (hasArtificialLightingSheet) {
                conditions.add("Отношение естественной освещенности к искусственной составляет");
            }
            return String.join(". ", conditions);
        } catch (Exception ex) {
            return "";
        }
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
        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            String name = readMergedCellValue(sheet, rowIndex, nameColumn, formatter);
            String serial = readMergedCellValue(sheet, rowIndex, serialColumn, formatter);
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

    private static String readMergedCellValue(Sheet sheet, int rowIndex, int colIndex, DataFormatter formatter) {
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
                return normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
            }
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell));
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
                return formatCellValue(cell, formatter, evaluator);
            }
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(colIndex);
        return formatCellValue(cell, formatter, evaluator);
    }

    private static Sheet findSheetWithName(Workbook workbook, String name) {
        if (workbook == null || name == null) {
            return null;
        }
        return workbook.getSheet(name);
    }

    private static String formatCellValue(Cell cell,
                                          DataFormatter formatter,
                                          FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
        try {
            String raw = (evaluator == null)
                    ? formatter.formatCellValue(cell)
                    : formatter.formatCellValue(cell, evaluator);
            return normalizeText(raw);
        } catch (Exception ex) {
            return normalizeText(formatter.formatCellValue(cell));
        }
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

    private static String findMeasurementPerformer(Sheet sheet, DataFormatter formatter) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                String rawText = formatter.formatCellValue(cell).trim();
                if (rawText.isEmpty()) {
                    continue;
                }
                String text = normalizeText(rawText);
                if (text.contains("Измерения проводил")) {
                    if (text.contains("Тарновский")) {
                        return "Тарновский М.О.";
                    }
                    if (text.contains("Белов")) {
                        return "Белов Д.А.";
                    }
                }
            }
        }
        return "";
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
    private static final String INSTRUMENTS_SECTION_HEADER = "сведения о средствах измерения:";
    private static final String INSTRUMENTS_NAME_HEADER = "наименование, тип средства измерения";
    private static final String PROTOCOL_PREFIX = "Протокол испытаний";
    private static final String BASIS_PREFIX = "Основание для измерений: договор";
    private static final java.util.regex.Pattern DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{2}\\.\\d{2}\\.\\d{4}\\b");
    private static final java.util.regex.Pattern CONTROL_DATE_PATTERN =
            java.util.regex.Pattern.compile("\\b\\d{1,2}\\s+(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s+\\d{4}\\b",
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
    private static boolean hasRowContentIncludingMerged(Sheet sheet,
                                                        int rowIndex,
                                                        int firstCol,
                                                        int lastCol,
                                                        DataFormatter formatter) {
        if (sheet == null) {
            return false;
        }
        for (int c = firstCol; c <= lastCol; c++) {
            String v = readMergedCellValue(sheet, rowIndex, c, formatter);
            if (!normalizeText(v).isBlank()) {
                return true;
            }
        }
        return false;
    }

}
