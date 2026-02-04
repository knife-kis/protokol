package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;

public final class HouseNoiseEquipmentIssuanceSheetExporter {
    private static final int NOISE_PROTOCOL_MERGED_DATE_LAST_COLUMN = 24;

    private HouseNoiseEquipmentIssuanceSheetExporter() {
    }

    public static void generate(File sourceNoiseProtocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }

        List<String> measurementDates = resolveNoiseMeasurementDatesFromProtocol(sourceNoiseProtocolFile);
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }

        String objectName = EquipmentIssuanceSheetExporter.resolveObjectName(mapFile);
        String performer = EquipmentIssuanceSheetExporter.resolveMeasurementPerformer(mapFile);
        List<EquipmentIssuanceSheetExporter.InstrumentEntry> instruments =
                EquipmentIssuanceSheetExporter.resolveInstruments(mapFile);

        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);
            File targetFile = resolveIssuanceSheetFileForNoise(mapFile, date, index, measurementDates.size());
            EquipmentIssuanceSheetExporter.writeIssuanceSheet(targetFile, objectName, performer, instruments, date);
        }
    }

    public static List<File> resolveIssuanceSheetFiles(File sourceNoiseProtocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }

        List<String> measurementDates = resolveNoiseMeasurementDatesFromProtocol(sourceNoiseProtocolFile);
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }

        List<File> files = new ArrayList<>();
        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);
            files.add(resolveIssuanceSheetFileForNoise(mapFile, date, index, measurementDates.size()));
        }
        return files;
    }

    private static List<String> resolveNoiseMeasurementDatesFromProtocol(File sourceNoiseProtocolFile) {
        if (sourceNoiseProtocolFile == null || !sourceNoiseProtocolFile.exists()) {
            return Collections.emptyList();
        }

        try (InputStream in = new FileInputStream(sourceNoiseProtocolFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() <= 1) {
                return Collections.emptyList();
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            Set<String> dates = new LinkedHashSet<>();

            for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                if (sheet == null || !isNoiseProtocolSheet(sheet.getSheetName())) {
                    continue;
                }

                List<CellRangeAddress> candidates = new ArrayList<>();
                for (CellRangeAddress region : sheet.getMergedRegions()) {
                    if (isNoiseProtocolDateRegion(region)) {
                        candidates.add(region);
                    }
                }
                candidates.sort(java.util.Comparator.comparingInt(CellRangeAddress::getFirstRow));

                for (CellRangeAddress region : candidates) {
                    String text = EquipmentIssuanceSheetExporter.readCellText(sheet,
                            region.getFirstRow(),
                            region.getFirstColumn(),
                            formatter,
                            evaluator);
                    if (!text.isEmpty()) {
                        EquipmentIssuanceSheetExporter.addDatesFromText(text, dates);
                    }
                }
            }

            return new ArrayList<>(dates);
        } catch (Exception ignored) {
            return Collections.emptyList();
        }
    }

    private static boolean isNoiseProtocolDateRegion(CellRangeAddress region) {
        return region.getFirstRow() == region.getLastRow()
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= NOISE_PROTOCOL_MERGED_DATE_LAST_COLUMN;
    }

    private static String normalizeText(String text) {
        if (text == null) {
            return "";
        }
        return text.replace('\u00A0', ' ').replaceAll("\\s+", " ").trim();
    }

    private static boolean isNoiseProtocolSheet(String sheetName) {
        String normalized = normalizeText(sheetName).toLowerCase(Locale.ROOT);
        return normalized.contains("шум");
    }

    private static File resolveIssuanceSheetFileForNoise(File mapFile, String date, int index, int total) {
        String name = "лист выдачи приборов";
        String safeDate = EquipmentIssuanceSheetExporter.sanitizeFileComponent(date);

        if (!safeDate.isBlank()) {
            name = name + " " + safeDate;
        } else if (total > 1) {
            name = name + " " + (index + 1);
        }

        return new File(mapFile.getParentFile(), name + ".docx");
    }
}
