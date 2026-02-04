package ru.citlab24.protokol.protocolmap.area;

import java.io.File;
import java.io.IOException;

final class AreaPrimaryRadiationExporter {
    private static final String PRIMARY_FOLDER_NAME = "Первичка Радиация (участки)";

    private AreaPrimaryRadiationExporter() {
    }

    static File generate(File sourceFile, String workDeadline, String customerInn) throws IOException {
        File mapFile = AreaRadiationMapExporter.generateMap(sourceFile, workDeadline, customerInn, PRIMARY_FOLDER_NAME);
        if (mapFile == null || !mapFile.exists()) {
            return mapFile;
        }
        SamplingPlanExporter.generate(sourceFile, mapFile);
        EquipmentControlSheetExporter.generate(mapFile);
        RadiationJournalExporter.generate(sourceFile, mapFile);
        RadiationJournalWordExporter.generate(sourceFile, mapFile);
        return mapFile;
    }
}
