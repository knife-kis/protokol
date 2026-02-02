package ru.citlab24.protokol.protocolmap;

import java.io.File;

final class SoundInsulationRequestFormExporter {
    private SoundInsulationRequestFormExporter() {
    }

    static void generate(File sourceFile, File mapFile, String workDeadline, String customerInn) {
        RequestFormExporter.generate(sourceFile, mapFile, workDeadline, customerInn);
    }

    static File resolveRequestFormFile(File mapFile) {
        return RequestFormExporter.resolveRequestFormFile(mapFile);
    }
}
