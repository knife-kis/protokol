package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.AreaNoiseEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.NoiseMapExporter;
import ru.citlab24.protokol.protocolmap.area.noise.AreaNoisePrimaryFiles;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

final class AreaPrimaryNoiseExporter {
    private static final String PRIMARY_FOLDER_NAME = "Первичка Шумы (участки)";
    private static final String AREA_NOISE_ADDITIONAL_INFO =
            "5. Дополнительные сведения: эскиз предоставлен заказчиком. При измерениях на объекте заказчик " +
                    "самостоятельно определяет расположение точек в соответствии с эскизом.";
    private static final String METHOD_HEADER =
            "Идентификация применяемого метода испытаний / метод (методика) измерений";

    private AreaPrimaryNoiseExporter() {
    }

    static File generate(File sourceFile, String workDeadline, String customerInn) throws IOException {
        File mapFile = generate(sourceFile, workDeadline, customerInn,
                ensurePrimaryFolder(sourceFile, PRIMARY_FOLDER_NAME), true);
        if (mapFile == null || !mapFile.exists()) {
            return mapFile;
        }

        File targetFolder = mapFile.getParentFile();
        List<File> filesToMove = AreaNoisePrimaryFiles.collect(mapFile, sourceFile);

        for (File file : filesToMove) {
            if (file == null || !file.exists()) {
                continue;
            }
            File target = new File(targetFolder, file.getName());
            if (file.equals(target)) {
                continue;
            }
            Files.move(file.toPath(), target.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }

        return mapFile;
    }

    static File generate(File sourceFile,
                         String workDeadline,
                         String customerInn,
                         File primaryFolder,
                         boolean generateCompanionDocuments) throws IOException {
        File mapFile = NoiseMapExporter.generateMap(sourceFile, workDeadline, customerInn,
                primaryFolder, generateCompanionDocuments);
        if (mapFile == null || !mapFile.exists()) {
            return mapFile;
        }
        applyAreaNoiseMapOverrides(sourceFile, mapFile);
        if (generateCompanionDocuments) {
            AreaNoiseEquipmentIssuanceSheetExporter.generate(sourceFile, mapFile);
        }
        return mapFile;
    }

    private static File ensurePrimaryFolder(File sourceFile, String folderName) {
        File root = sourceFile != null ? sourceFile.getParentFile() : null;
        if (root == null) {
            root = new File(".");
        }
        File targetFolder = new File(root, folderName);
        if (!targetFolder.exists()) {
            targetFolder.mkdirs();
        }
        return targetFolder;
    }

    private static void applyAreaNoiseMapOverrides(File sourceFile, File mapFile) {
        String method = resolveAreaNoiseMethod(sourceFile);
        try (InputStream inputStream = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            DataFormatter formatter = new DataFormatter();
            for (Sheet sheet : workbook) {
                if (!hasAreaNoiseTable(sheet, formatter)) {
                    continue;
                }
                for (Row row : sheet) {
                    Cell firstCell = row.getCell(0);
                    if (firstCell == null) {
                        continue;
                    }
                    String normalized = normalize(formatter.formatCellValue(firstCell)).toLowerCase(Locale.ROOT);
                    if (normalized.startsWith("5. дополнительные сведения")) {
                        firstCell.setCellValue(AREA_NOISE_ADDITIONAL_INFO);
                    } else if (!method.isBlank() && normalized.startsWith("5.2. методы измерения")) {
                        firstCell.setCellValue("5.2. Методы измерения " + ensureTrailingSemicolon(method));
                        row.setHeightInPoints(row.getHeightInPoints() + sheet.getDefaultRowHeightInPoints());
                    }
                }
                mergeAreaNoisePlaceCells(sheet, formatter);
            }
            try (FileOutputStream outputStream = new FileOutputStream(mapFile)) {
                workbook.write(outputStream);
            }
        } catch (Exception ignored) {
            // карта уже сформирована, пропускаем только точечную правку текста
        }
    }

    private static String resolveAreaNoiseMethod(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream inputStream = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(inputStream)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String normalized = normalize(formatter.formatCellValue(cell, evaluator));
                    if (!normalized.equalsIgnoreCase(METHOD_HEADER)) {
                        continue;
                    }
                    return collectMethodsBelow(sheet, row.getRowNum(), cell.getColumnIndex(), formatter, evaluator);
                }
            }
        } catch (Exception ignored) {
            return "";
        }
        return "";
    }

    private static String collectMethodsBelow(Sheet sheet,
                                              int headerRowIndex,
                                              int columnIndex,
                                              DataFormatter formatter,
                                              FormulaEvaluator evaluator) {
        Set<String> methods = new LinkedHashSet<>();
        int emptyStreak = 0;
        for (int rowIndex = headerRowIndex + 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            String value = readMergedCellValue(sheet, rowIndex, columnIndex, formatter, evaluator);
            String normalized = normalize(value);
            if (normalized.isBlank()) {
                emptyStreak++;
                if (emptyStreak >= 2) {
                    break;
                }
                continue;
            }
            emptyStreak = 0;
            if (isMethodSectionEnd(normalized)) {
                break;
            }
            methods.add(normalized);
        }
        return String.join("; ", new ArrayList<>(methods));
    }

    private static boolean hasAreaNoiseTable(Sheet sheet, DataFormatter formatter) {
        if (sheet == null) {
            return false;
        }
        for (Row row : sheet) {
            Cell firstCell = row.getCell(0);
            String value = normalize(formatCellValue(firstCell, formatter)).toLowerCase(Locale.ROOT);
            if (value.startsWith("7.6.2. шум") || value.startsWith("5. дополнительные сведения")) {
                return true;
            }
        }
        return false;
    }

    private static void mergeAreaNoisePlaceCells(Sheet sheet, DataFormatter formatter) {
        for (int rowIndex = sheet.getFirstRowNum(); rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            Cell placeCell = row == null ? null : row.getCell(2);
            String place = normalize(formatCellValue(placeCell, formatter)).toLowerCase(Locale.ROOT);
            if (!"земельный участок".equals(place) || findMergedRegion(sheet, rowIndex, 2) != null) {
                continue;
            }
            int endRow = rowIndex;
            while (endRow + 1 <= sheet.getLastRowNum()) {
                Row nextRow = sheet.getRow(endRow + 1);
                if (!hasNoisePointRow(nextRow, formatter)
                        || !normalize(formatCellValue(nextRow == null ? null : nextRow.getCell(2), formatter)).isBlank()) {
                    break;
                }
                endRow++;
            }
            if (endRow > rowIndex) {
                CellRangeAddress region = new CellRangeAddress(rowIndex, endRow, 2, 2);
                sheet.addMergedRegion(region);
                applyMergedRegionBorders(region, sheet, placeCell.getCellStyle());
            }
        }
    }

    private static boolean hasNoisePointRow(Row row, DataFormatter formatter) {
        return !normalize(formatCellValue(row == null ? null : row.getCell(0), formatter)).isBlank()
                || !normalize(formatCellValue(row == null ? null : row.getCell(1), formatter)).isBlank();
    }

    private static String formatCellValue(Cell cell, DataFormatter formatter) {
        return cell == null ? "" : formatter.formatCellValue(cell);
    }

    private static void applyMergedRegionBorders(CellRangeAddress region, Sheet sheet, CellStyle style) {
        RegionUtil.setBorderTop(style.getBorderTop(), region, sheet);
        RegionUtil.setBorderBottom(style.getBorderBottom(), region, sheet);
        RegionUtil.setBorderLeft(style.getBorderLeft(), region, sheet);
        RegionUtil.setBorderRight(style.getBorderRight(), region, sheet);
    }

    private static boolean isMethodSectionEnd(String value) {
        String lower = normalize(value).toLowerCase(Locale.ROOT);
        return lower.matches("^\\d+\\.\\s+.*")
                || lower.startsWith("дополнительные сведения")
                || lower.startsWith("сведения о средствах измерения")
                || lower.startsWith("особые условия");
    }

    private static String ensureTrailingSemicolon(String value) {
        String trimmed = normalize(value);
        if (trimmed.isBlank() || trimmed.endsWith(";")) {
            return trimmed;
        }
        return trimmed + ";";
    }

    private static String readMergedCellValue(Sheet sheet,
                                              int rowIndex,
                                              int columnIndex,
                                              DataFormatter formatter,
                                              FormulaEvaluator evaluator) {
        CellRangeAddress region = findMergedRegion(sheet, rowIndex, columnIndex);
        int targetRow = region == null ? rowIndex : region.getFirstRow();
        int targetCol = region == null ? columnIndex : region.getFirstColumn();
        Row row = sheet.getRow(targetRow);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(targetCol);
        return cell == null ? "" : formatter.formatCellValue(cell, evaluator);
    }

    private static CellRangeAddress findMergedRegion(Sheet sheet, int rowIndex, int columnIndex) {
        for (CellRangeAddress region : sheet.getMergedRegions()) {
            if (region.isInRange(rowIndex, columnIndex)) {
                return region;
            }
        }
        return null;
    }

    private static String normalize(String value) {
        if (value == null) {
            return "";
        }
        return value.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }

}
