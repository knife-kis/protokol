package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.NoiseMapExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

final class AreaPrimaryNoiseExporter {
    private static final String PRIMARY_FOLDER_NAME = "Первичка Шумы (участки)";

    private AreaPrimaryNoiseExporter() {
    }

    static File generate(File sourceFile, String workDeadline, String customerInn) throws IOException {
        File mapFile = NoiseMapExporter.generateMap(sourceFile, workDeadline, customerInn);
        if (mapFile == null || !mapFile.exists()) {
            return mapFile;
        }

        File requestForm = RequestFormExporter.resolveRequestFormFile(mapFile);
        deleteIfExists(requestForm);
        File registrationSheet = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mapFile);
        deleteIfExists(registrationSheet);

        File targetFolder = ensurePrimaryFolder(mapFile);
        List<File> filesToMove = new ArrayList<>();
        filesToMove.add(mapFile);
        File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(mapFile);
        if (analysisSheet != null) {
            filesToMove.add(analysisSheet);
        }
        File measurementPlan = MeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
        if (measurementPlan != null) {
            filesToMove.add(measurementPlan);
        }
        File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
        if (issuanceSheet != null) {
            filesToMove.add(issuanceSheet);
        }
        filesToMove.addAll(EquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(mapFile));

        for (File file : filesToMove) {
            if (file == null || !file.exists()) {
                continue;
            }
            File target = new File(targetFolder, file.getName());
            Files.move(file.toPath(), target.toPath(), StandardCopyOption.REPLACE_EXISTING);
        }

        File newMapFile = new File(targetFolder, mapFile.getName());
        if (!newMapFile.exists()) {
            return mapFile;
        }
        ShortRequestFormExporter.generate(newMapFile);
        return newMapFile;
    }

    private static void deleteIfExists(File file) throws IOException {
        if (file == null) {
            return;
        }
        Files.deleteIfExists(file.toPath());
    }

    private static File ensurePrimaryFolder(File mapFile) {
        File currentFolder = mapFile != null ? mapFile.getParentFile() : null;
        File root = currentFolder != null ? currentFolder.getParentFile() : new File(".");
        File targetFolder = new File(root, PRIMARY_FOLDER_NAME);
        if (!targetFolder.exists()) {
            targetFolder.mkdirs();
        }
        return targetFolder;
    }
}
