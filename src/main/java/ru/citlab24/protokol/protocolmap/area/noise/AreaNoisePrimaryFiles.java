package ru.citlab24.protokol.protocolmap.area.noise;

import ru.citlab24.protokol.protocolmap.AreaNoiseEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

public final class AreaNoisePrimaryFiles {
    private AreaNoisePrimaryFiles() {
    }

    public static List<File> collect(File mapFile, File sourceNoiseProtocolFile) {
        List<File> files = new ArrayList<>();
        if (mapFile != null) {
            files.add(mapFile);
        }
        File requestForm = RequestFormExporter.resolveRequestFormFile(mapFile);
        if (requestForm != null) {
            files.add(requestForm);
        }
        File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(mapFile);
        if (analysisSheet != null) {
            files.add(analysisSheet);
        }
        File measurementPlan = MeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
        if (measurementPlan != null) {
            files.add(measurementPlan);
        }
        File registrationSheet = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mapFile);
        if (registrationSheet != null) {
            files.add(registrationSheet);
        }
        List<File> equipmentSheets = AreaNoiseEquipmentIssuanceSheetExporter
                .resolveIssuanceSheetFiles(sourceNoiseProtocolFile, mapFile);
        if (equipmentSheets != null) {
            files.addAll(equipmentSheets);
        }
        File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
        if (issuanceSheet != null) {
            files.add(issuanceSheet);
        }
        return files;
    }

    public static File resolveRequestFormFile(File mapFile) {
        return RequestFormExporter.resolveRequestFormFile(mapFile);
    }

    public static File resolveAnalysisSheetFile(File mapFile) {
        return RequestAnalysisSheetExporter.resolveAnalysisSheetFile(mapFile);
    }

    public static File resolveMeasurementPlanFile(File mapFile) {
        return MeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
    }

    public static File resolveRegistrationSheetFile(File mapFile) {
        return MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mapFile);
    }

    public static List<File> resolveEquipmentIssuanceFiles(File sourceNoiseProtocolFile, File mapFile) {
        return AreaNoiseEquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(sourceNoiseProtocolFile, mapFile);
    }

    public static File resolveProtocolIssuanceSheetFile(File mapFile) {
        return ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
    }
}
