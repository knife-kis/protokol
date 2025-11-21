package ru.citlab24.protokol.db;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;
import ru.citlab24.protokol.tabs.models.Section;
import ru.citlab24.protokol.tabs.models.Space;

import javax.swing.*;
import java.awt.*;
import java.io.*;
import java.sql.SQLException;
import java.util.HashMap;
import java.util.Map;

/**
 * Сервис для обмена проектами через файл.
 */
public final class ProjectFileService {
    private static final Logger logger = LoggerFactory.getLogger(ProjectFileService.class);
    private static final String DEFAULT_EXTENSION = "protokol";

    private ProjectFileService() {}

    /**
     * Выгружает сохранённый проект в файл. Если buildingId <= 0 — выводит информационное сообщение и ничего не делает.
     */
    public static void exportProject(Component parent, int buildingId, String projectName)
            throws SQLException, IOException {
        if (buildingId <= 0) {
            JOptionPane.showMessageDialog(parent,
                    "Сначала сохраните проект, чтобы можно было экспортировать файл.",
                    "Экспорт проекта", JOptionPane.INFORMATION_MESSAGE);
            return;
        }

        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Экспорт проекта");
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Файл проекта", DEFAULT_EXTENSION
        ));

        String baseName = (projectName == null || projectName.isBlank())
                ? "project"
                : projectName.trim();
        baseName = baseName.replaceAll("[\\\\/:*?\"<>|]", "_");
        chooser.setSelectedFile(new File(baseName + "." + DEFAULT_EXTENSION));

        int result = chooser.showSaveDialog(parent);
        if (result != JFileChooser.APPROVE_OPTION) return;

        File target = ensureExtension(chooser.getSelectedFile());
        ProjectSnapshot snapshot = buildSnapshot(buildingId);

        try (ObjectOutputStream oos = new ObjectOutputStream(new BufferedOutputStream(new FileOutputStream(target)))) {
            oos.writeObject(snapshot);
        }

        logger.info("Проект {} (id={}) экспортирован в {}", projectName, buildingId, target.getAbsolutePath());

        JOptionPane.showMessageDialog(parent,
                "Проект сохранён в файл:\n" + target.getAbsolutePath(),
                "Экспорт проекта", JOptionPane.INFORMATION_MESSAGE);
    }

    /**
     * Импортирует проект из файла и сохраняет его в локальную базу. Возвращает загруженное здание для отображения в UI.
     */
    public static Building importProject(Component parent)
            throws IOException, ClassNotFoundException, SQLException {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Импорт проекта");
        chooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
                "Файл проекта", DEFAULT_EXTENSION
        ));

        int result = chooser.showOpenDialog(parent);
        if (result != JFileChooser.APPROVE_OPTION) return null;

        File source = chooser.getSelectedFile();
        logger.info("Импорт проекта из файла {}", source.getAbsolutePath());
        ProjectSnapshot snapshot;
        try (ObjectInputStream ois = new ObjectInputStream(new BufferedInputStream(new FileInputStream(source)))) {
            snapshot = (ProjectSnapshot) ois.readObject();
        }

        if (snapshot == null || snapshot.building == null) {
            JOptionPane.showMessageDialog(parent,
                    "Файл проекта повреждён или пуст", "Импорт", JOptionPane.ERROR_MESSAGE);
            return null;
        }

        Building saved = persistSnapshot(snapshot);
        JOptionPane.showMessageDialog(parent,
                "Проект успешно импортирован:\n" + saved.getName(),
                "Импорт проекта", JOptionPane.INFORMATION_MESSAGE);
        return saved;
    }

    private static ProjectSnapshot buildSnapshot(int buildingId) throws SQLException {
        Building building = DatabaseManager.loadBuilding(buildingId);
        Map<String, Double[]> streetLighting = DatabaseManager.loadStreetLightingValuesByKey(buildingId);
        Map<String, DatabaseManager.NoiseValue> noiseSelections = DatabaseManager.loadNoiseSelectionsByKey(buildingId);
        Map<String, double[]> thresholds = DatabaseManager.loadNoiseThresholds(buildingId);
        return new ProjectSnapshot(building, streetLighting, noiseSelections, thresholds);
    }

    private static Building persistSnapshot(ProjectSnapshot snapshot) throws SQLException {
        Building building = snapshot.building;
        normalizeIds(building);
        DatabaseManager.saveBuilding(building);

        if (!snapshot.streetLightingValues.isEmpty()) {
            DatabaseManager.updateStreetLightingValues(building, snapshot.streetLightingValues);
        }
        if (!snapshot.noiseSelections.isEmpty()) {
            DatabaseManager.updateNoiseSelections(building, snapshot.noiseSelections);
        }
        if (!snapshot.noiseThresholds.isEmpty()) {
            DatabaseManager.updateNoiseThresholds(building, snapshot.noiseThresholds);
        }

        return DatabaseManager.loadBuilding(building.getId());
    }

    private static void normalizeIds(Building building) {
        if (building == null) return;
        building.setId(0);

        int sectionPos = 0;
        for (Section section : building.getSections()) {
            if (section == null) continue;
            section.setId(0);
            section.setPosition(sectionPos++);
        }

        for (Floor floor : building.getFloors()) {
            if (floor == null) continue;
            floor.setId(0);
            floor.setPosition(Math.max(0, floor.getPosition()));
            floor.setSectionIndex(Math.max(0, floor.getSectionIndex()));

            for (Space space : floor.getSpaces()) {
                if (space == null) continue;
                space.setId(0);
                space.setPosition(Math.max(0, space.getPosition()));

                for (Room room : space.getRooms()) {
                    if (room == null) continue;
                    Integer originalId = room.getOriginalRoomId();
                    if (originalId == null || originalId <= 0) {
                        room.setOriginalRoomId(room.getId());
                    }
                    room.setId(0);
                    room.setPosition(Math.max(0, room.getPosition()));
                }
            }
        }
    }

    private static File ensureExtension(File file) {
        if (file == null) return null;
        String name = file.getName();
        if (name.toLowerCase().endsWith("." + DEFAULT_EXTENSION)) return file;
        return new File(file.getParentFile(), name + "." + DEFAULT_EXTENSION);
    }

    private static final class ProjectSnapshot implements Serializable {
        private static final long serialVersionUID = 1L;
        private final Building building;
        private final Map<String, Double[]> streetLightingValues;
        private final Map<String, DatabaseManager.NoiseValue> noiseSelections;
        private final Map<String, double[]> noiseThresholds;

        private ProjectSnapshot(Building building,
                                Map<String, Double[]> streetLightingValues,
                                Map<String, DatabaseManager.NoiseValue> noiseSelections,
                                Map<String, double[]> noiseThresholds) {
            this.building = building;
            this.streetLightingValues = (streetLightingValues == null)
                    ? new HashMap<>() : new HashMap<>(streetLightingValues);
            this.noiseSelections = (noiseSelections == null)
                    ? new HashMap<>() : new HashMap<>(noiseSelections);
            this.noiseThresholds = (noiseThresholds == null)
                    ? new HashMap<>() : new HashMap<>(noiseThresholds);
        }
    }
}
