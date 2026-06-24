package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;
import ru.citlab24.protokol.protocolmap.AreaNoiseEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

final class AreaPrimaryBatchExporter {
    static final String PRIMARY_FOLDER_NAME = "первичка участки";

    private AreaPrimaryBatchExporter() {
    }

    static Result generate(File radiationSource,
                           File noiseSource,
                           String workDeadline,
                           String customerInn) throws IOException {
        List<ModuleDocument> documents = new ArrayList<>();
        File primaryFolder = resolvePrimaryFolder(radiationSource != null ? radiationSource : noiseSource);

        if (radiationSource != null && radiationSource.exists()) {
            File map = AreaRadiationMapExporter.generateMap(radiationSource, workDeadline, customerInn,
                    primaryFolder, false);
            documents.add(new ModuleDocument(ModuleKind.RADIATION, radiationSource, map));
        }
        if (noiseSource != null && noiseSource.exists()) {
            File map = AreaPrimaryNoiseExporter.generate(noiseSource, workDeadline, customerInn,
                    primaryFolder, false);
            documents.add(new ModuleDocument(ModuleKind.NOISE, noiseSource, map));
        }

        if (documents.isEmpty()) {
            return new Result(primaryFolder, documents);
        }

        generateModuleSpecificDocuments(documents, workDeadline);
        generateCombinedProtocolIssuance(documents);
        generateCombinedMeasurementRegistration(documents);
        generateRequestForm(documents, workDeadline, customerInn);
        generateRequestAnalysis(documents);

        return new Result(primaryFolder, documents);
    }

    private static File resolvePrimaryFolder(File sourceFile) {
        File root = sourceFile != null ? sourceFile.getParentFile() : null;
        if (root == null) {
            root = new File(".");
        }
        File folder = new File(root, PRIMARY_FOLDER_NAME);
        if (!folder.exists()) {
            folder.mkdirs();
        }
        return folder;
    }

    private static void generateModuleSpecificDocuments(List<ModuleDocument> documents, String workDeadline)
            throws IOException {
        for (ModuleDocument document : documents) {
            if (document.mapFile == null || !document.mapFile.exists()) {
                continue;
            }
            if (document.kind == ModuleKind.RADIATION) {
                EquipmentIssuanceSheetExporter.generate(document.mapFile);
                MeasurementPlanExporter.generate(document.sourceFile, document.mapFile, workDeadline);
                SamplingPlanExporter.generate(document.sourceFile, document.mapFile);
                EquipmentControlSheetExporter.generate(document.mapFile);
                RadiationJournalWordExporter.generate(document.sourceFile, document.mapFile);
            } else {
                AreaNoiseEquipmentIssuanceSheetExporter.generate(document.sourceFile, document.mapFile);
                MeasurementPlanExporter.generateForNoise(document.sourceFile, document.mapFile, workDeadline);
            }
        }
    }

    private static void generateCombinedProtocolIssuance(List<ModuleDocument> documents) {
        ModuleDocument main = documents.get(0);
        ProtocolIssuanceSheetExporter.generateCombined(
                ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(main.mapFile),
                sourceFiles(documents),
                mapFiles(documents));
    }

    private static void generateCombinedMeasurementRegistration(List<ModuleDocument> documents) {
        ModuleDocument main = documents.get(0);
        MeasurementCardRegistrationSheetExporter.generateCombined(
                MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(main.mapFile),
                sourceFiles(documents),
                mapFiles(documents));
    }

    private static void generateRequestForm(List<ModuleDocument> documents,
                                            String workDeadline,
                                            String customerInn) {
        ModuleDocument main = documents.get(0);
        File target = generateSingleRequestForm(main, workDeadline, customerInn, hasModule(documents, ModuleKind.NOISE));
        if (target == null || !target.exists()) {
            return;
        }
        for (int i = 1; i < documents.size(); i++) {
            TemporaryRequestForm extra = generateTemporaryRequestForm(documents.get(i), workDeadline, customerInn);
            appendDocument(target, extra.file);
            deleteTemporaryRequestForm(extra);
        }
    }

    private static File generateSingleRequestForm(ModuleDocument document,
                                                  String workDeadline,
                                                  String customerInn,
                                                  boolean includeNoiseInstructions) {
        if (document.kind == ModuleKind.NOISE) {
            RequestFormExporter.generateForNoise(document.sourceFile, document.mapFile, workDeadline, customerInn);
        } else {
            RequestFormExporter.generateAreaRadiation(document.sourceFile, document.mapFile, workDeadline,
                    customerInn, includeNoiseInstructions);
        }
        return RequestFormExporter.resolveRequestFormFile(document.mapFile);
    }

    private static TemporaryRequestForm generateTemporaryRequestForm(ModuleDocument document,
                                                                     String workDeadline,
                                                                     String customerInn) {
        if (document.mapFile == null || !document.mapFile.exists()) {
            return TemporaryRequestForm.empty();
        }
        File tempDirectory = null;
        try {
            tempDirectory = Files.createTempDirectory("area-primary-request-").toFile();
            File tempMap = new File(tempDirectory, document.mapFile.getName());
            Files.copy(document.mapFile.toPath(), tempMap.toPath());
            if (document.kind == ModuleKind.NOISE) {
                RequestFormExporter.generateForNoiseAppendix(document.sourceFile, tempMap, workDeadline, customerInn);
            } else {
                RequestFormExporter.generate(document.sourceFile, tempMap, workDeadline, customerInn);
            }
            return new TemporaryRequestForm(RequestFormExporter.resolveRequestFormFile(tempMap), tempDirectory);
        } catch (Exception ignored) {
            deleteDirectory(tempDirectory);
            return TemporaryRequestForm.empty();
        }
    }

    private static void generateRequestAnalysis(List<ModuleDocument> documents) {
        ModuleDocument main = documents.get(0);
        RequestAnalysisSheetExporter.generateForNoise(main.mapFile, collectAdditionalMeasurementDates(documents));
    }

    private static boolean hasModule(List<ModuleDocument> documents, ModuleKind kind) {
        for (ModuleDocument document : documents) {
            if (document.kind == kind) {
                return true;
            }
        }
        return false;
    }

    private static List<String> collectAdditionalMeasurementDates(List<ModuleDocument> documents) {
        List<String> dates = new ArrayList<>();
        for (int i = 1; i < documents.size(); i++) {
            dates.addAll(RequestAnalysisSheetExporter.resolveMeasurementDatesForMap(documents.get(i).mapFile));
        }
        return dates;
    }

    private static List<File> sourceFiles(List<ModuleDocument> documents) {
        List<File> files = new ArrayList<>();
        for (ModuleDocument document : documents) {
            files.add(document.sourceFile);
        }
        return files;
    }

    private static List<File> mapFiles(List<ModuleDocument> documents) {
        List<File> files = new ArrayList<>();
        for (ModuleDocument document : documents) {
            files.add(document.mapFile);
        }
        return files;
    }

    private static void appendDocument(File targetFile, File sourceFile) {
        if (targetFile == null || sourceFile == null || !targetFile.exists() || !sourceFile.exists()
                || targetFile.equals(sourceFile)) {
            return;
        }
        XWPFDocument target;
        try (InputStream targetInput = new FileInputStream(targetFile)) {
            target = new XWPFDocument(targetInput);
        } catch (Exception ignored) {
            return;
        }
        try (XWPFDocument targetDocument = target;
             InputStream sourceInput = new FileInputStream(sourceFile);
             XWPFDocument source = new XWPFDocument(sourceInput)) {
            XWPFParagraph breakParagraph = targetDocument.createParagraph();
            breakParagraph.createRun().addBreak(BreakType.PAGE);
            for (IBodyElement element : source.getBodyElements()) {
                XmlCursor cursor = targetDocument.getDocument().getBody().newCursor();
                cursor.toEndToken();
                if (element instanceof XWPFParagraph paragraph) {
                    XWPFParagraph targetParagraph = targetDocument.insertNewParagraph(cursor);
                    targetParagraph.getCTP().set(paragraph.getCTP().copy());
                } else if (element instanceof XWPFTable table) {
                    XWPFTable targetTable = targetDocument.insertNewTbl(cursor);
                    targetTable.getCTTbl().set(table.getCTTbl().copy());
                }
                cursor.dispose();
            }
            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
                targetDocument.write(out);
                Files.write(targetFile.toPath(), out.toByteArray());
            }
        } catch (Exception ignored) {
        }
    }

    private static void deleteTemporaryRequestForm(TemporaryRequestForm requestForm) {
        if (requestForm != null) {
            deleteDirectory(requestForm.directory);
        }
    }

    private static void deleteDirectory(File directory) {
        if (directory == null || !directory.exists()) {
            return;
        }
        File[] files = directory.listFiles();
        if (files != null) {
            for (File file : files) {
                if (file.isDirectory()) {
                    deleteDirectory(file);
                } else {
                    file.delete();
                }
            }
        }
        directory.delete();
    }

    enum ModuleKind {
        RADIATION,
        NOISE
    }

    static final class ModuleDocument {
        final ModuleKind kind;
        final File sourceFile;
        final File mapFile;

        ModuleDocument(ModuleKind kind, File sourceFile, File mapFile) {
            this.kind = kind;
            this.sourceFile = sourceFile;
            this.mapFile = mapFile;
        }
    }

    static final class Result {
        final File primaryFolder;
        final List<ModuleDocument> documents;

        Result(File primaryFolder, List<ModuleDocument> documents) {
            this.primaryFolder = primaryFolder;
            this.documents = documents;
        }
    }

    private static final class TemporaryRequestForm {
        private final File file;
        private final File directory;

        private TemporaryRequestForm(File file, File directory) {
            this.file = file;
            this.directory = directory;
        }

        private static TemporaryRequestForm empty() {
            return new TemporaryRequestForm(null, null);
        }
    }
}
