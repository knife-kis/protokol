package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.protocolmap.PhysicalFactorsMapExporter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

final class AreaRadiationMapExporter {
    private AreaRadiationMapExporter() {
    }

    static File generateMap(File sourceFile,
                            String workDeadline,
                            String customerInn,
                            String primaryFolderName) throws java.io.IOException {
        File mapFile = PhysicalFactorsMapExporter.generateMap(sourceFile, workDeadline, customerInn, primaryFolderName);
        if (mapFile == null || !mapFile.exists()) {
            return mapFile;
        }
        updateRadiationMap(sourceFile, mapFile);
        return mapFile;
    }

    private static void updateRadiationMap(File sourceFile, File mapFile) {
        if (sourceFile == null || !sourceFile.exists() || mapFile == null || !mapFile.exists()) {
            return;
        }
        boolean hasMedSheet = hasSheetWithName(sourceFile, "МЭД");
        int[] counts = hasMedSheet ? resolveMedProfileAndControlPointCounts(sourceFile) : new int[]{0, 0};

        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            removeSheetIfExists(workbook, "Вентиляция");
            if (hasMedSheet) {
                removeSheetIfExists(workbook, "МЭД");
                Sheet medSheet = AreaRadiationMedMapTabBuilder.createSheet(workbook, counts[0], counts[1]);
                applyHeadersFromMainSheet(workbook, medSheet);
            }
            try (FileOutputStream out = new FileOutputStream(mapFile)) {
                workbook.write(out);
            }
        } catch (Exception ex) {
            // ignore
        }
    }

    private static void removeSheetIfExists(Workbook workbook, String name) {
        if (workbook == null || name == null) {
            return;
        }
        int index = workbook.getSheetIndex(name);
        if (index >= 0) {
            workbook.removeSheetAt(index);
        }
    }

    private static void applyHeadersFromMainSheet(Workbook workbook, Sheet targetSheet) {
        if (workbook == null || targetSheet == null) {
            return;
        }
        Sheet sourceSheet = workbook.getSheet("карта замеров");
        if (sourceSheet == null) {
            sourceSheet = workbook.getNumberOfSheets() > 0 ? workbook.getSheetAt(0) : null;
        }
        if (sourceSheet == null) {
            return;
        }
        Header sourceHeader = sourceSheet.getHeader();
        Header targetHeader = targetSheet.getHeader();
        targetHeader.setLeft(sourceHeader.getLeft());
        targetHeader.setCenter(sourceHeader.getCenter());
        targetHeader.setRight(sourceHeader.getRight());
    }

    private static int[] resolveMedProfileAndControlPointCounts(File sourceFile) {
        int maxProfile = 0;
        int maxControlPoint = 0;
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook sourceWorkbook = WorkbookFactory.create(in)) {
            Sheet sourceSheet = findSheetWithName(sourceWorkbook, "МЭД");
            if (sourceSheet == null) {
                return new int[]{0, 0};
            }
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = sourceWorkbook.getCreationHelper().createFormulaEvaluator();
            Pattern profilePattern = Pattern.compile("(?i)профиль\\s*(\\d+)");
            Pattern controlPattern = Pattern.compile("(?i)контрольная\\s*точка\\s*(\\d+)");
            int lastRow = sourceSheet.getLastRowNum();
            int emptyStreak = 0;
            for (int rowIndex = 6; rowIndex <= lastRow; rowIndex++) {
                String bValue = readMergedCellValue(sourceSheet, rowIndex, 1, formatter, evaluator);
                String normalized = normalizeText(bValue);
                if (normalized.isBlank()) {
                    emptyStreak++;
                    if (emptyStreak >= 20) {
                        break;
                    }
                    continue;
                }
                emptyStreak = 0;
                Matcher profileMatcher = profilePattern.matcher(normalized);
                if (profileMatcher.find()) {
                    maxProfile = Math.max(maxProfile, Integer.parseInt(profileMatcher.group(1)));
                }
                Matcher controlMatcher = controlPattern.matcher(normalized);
                if (controlMatcher.find()) {
                    maxControlPoint = Math.max(maxControlPoint, Integer.parseInt(controlMatcher.group(1)));
                }
            }
        } catch (Exception ex) {
            return new int[]{0, 0};
        }
        if (maxProfile <= 0) {
            maxProfile = 1;
        }
        return new int[]{maxProfile, maxControlPoint};
    }

    private static Sheet findSheetWithName(Workbook workbook, String name) {
        if (workbook == null || name == null) {
            return null;
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (name.equalsIgnoreCase(normalizeText(sheet.getSheetName()))) {
                return sheet;
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
        CellRangeAddress region = findMergedRegion(sheet, rowIndex, colIndex);
        if (region != null) {
            Row row = sheet.getRow(region.getFirstRow());
            if (row == null) {
                return "";
            }
            org.apache.poi.ss.usermodel.Cell cell = row.getCell(region.getFirstColumn());
            return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        org.apache.poi.ss.usermodel.Cell cell = row.getCell(colIndex);
        return normalizeText(cell == null ? "" : formatter.formatCellValue(cell, evaluator));
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int colIndex) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, colIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String normalizeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }
}
