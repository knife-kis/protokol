package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.AreaNoiseEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.NoiseMapExporter;
import ru.citlab24.protokol.protocolmap.area.noise.AreaNoisePrimaryFiles;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
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

        AreaNoiseEquipmentIssuanceSheetExporter.generate(sourceFile, mapFile);

        File targetFolder = ensurePrimaryFolder(mapFile);
        List<File> filesToMove = AreaNoisePrimaryFiles.collect(mapFile, sourceFile);

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
        return newMapFile;
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
