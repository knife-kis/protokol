package ru.citlab24.protokol.protocolmap;

import java.io.File;
import java.io.IOException;

public final class SoundInsulationMapExporter {
    private static final String PRIMARY_FOLDER_NAME = "Первичка Звукоизоляция";

    private SoundInsulationMapExporter() {
    }

    public static File generateMap(File sourceFile, String workDeadline, String customerInn) throws IOException {
        return PhysicalFactorsMapExporter.generateMap(sourceFile, workDeadline, customerInn, PRIMARY_FOLDER_NAME);
    }
}
