package ru.citlab24.protokol.protocolmap;

import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

public final class HousePrimaryBatchExporter {
    public enum ModuleKind {
        PHYSICAL,
        NOISE,
        SOUND_INSULATION
    }

    public static final class ModuleDocument {
        public final ModuleKind kind;
        public final File sourceFile;
        public final File mapFile;
        public final File protocolFile;

        public ModuleDocument(ModuleKind kind, File sourceFile, File mapFile, File protocolFile) {
            this.kind = kind;
            this.sourceFile = sourceFile;
            this.mapFile = mapFile;
            this.protocolFile = protocolFile;
        }
    }

    private HousePrimaryBatchExporter() {
    }

    public static void generateCompanionDocuments(List<ModuleDocument> documents,
                                                  String workDeadline,
                                                  String customerInn) throws IOException {
        if (documents == null || documents.isEmpty()) {
            return;
        }
        ModuleDocument main = documents.get(0);
        File mainMap = main.mapFile;
        if (mainMap == null || !mainMap.exists()) {
            return;
        }

        generateModuleSpecificDocuments(documents, workDeadline);
        generateCombinedProtocolIssuance(mainMap, documents);
        generateCombinedMeasurementRegistration(mainMap, documents);
        generateRequestForm(main, documents, workDeadline, customerInn);
        generateRequestAnalysis(main, documents);
    }

    private static void generateModuleSpecificDocuments(List<ModuleDocument> documents, String workDeadline)
            throws IOException {
        for (ModuleDocument document : documents) {
            if (document.mapFile == null || !document.mapFile.exists()) {
                continue;
            }
            if (document.kind == ModuleKind.PHYSICAL) {
                EquipmentIssuanceSheetExporter.generate(document.mapFile);
                MeasurementPlanExporter.generate(document.sourceFile, document.mapFile, workDeadline);
                PhysicalFactorsValuesGuideExporter.generate(document.sourceFile, document.mapFile);
            } else if (document.kind == ModuleKind.NOISE) {
                HouseNoiseEquipmentIssuanceSheetExporter.generate(document.sourceFile, document.mapFile);
                MeasurementPlanExporter.generateForNoise(document.sourceFile, document.mapFile, workDeadline);
            } else if (document.kind == ModuleKind.SOUND_INSULATION) {
                SoundInsulationEquipmentIssuanceSheetExporter.generate(document.protocolFile, document.mapFile);
                SoundInsulationMeasurementPlanExporter.generate(document.protocolFile, document.mapFile, workDeadline);
            }
        }
    }

    private static void generateCombinedProtocolIssuance(File mainMap, List<ModuleDocument> documents) {
        List<ProtocolIssuanceSheetExporter.RowData> rows = new ArrayList<>();
        for (ModuleDocument document : documents) {
            if (document.kind == ModuleKind.SOUND_INSULATION) {
                SoundInsulationProtocolDataParser.ProtocolData data =
                        SoundInsulationProtocolDataParser.parse(document.protocolFile);
                String applicationNumber = data.applicationNumber();
                if (applicationNumber.isBlank()) {
                    applicationNumber = data.registrationNumber();
                }
                rows.add(new ProtocolIssuanceSheetExporter.RowData(
                        data.protocolNumber(),
                        data.protocolDate(),
                        data.customerName(),
                        applicationNumber
                ));
            } else {
                rows.add(ProtocolIssuanceSheetExporter.rowDataFromMap(document.sourceFile, document.mapFile));
            }
        }
        ProtocolIssuanceSheetExporter.generateCombinedRows(
                ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mainMap), rows);
    }

    private static void generateCombinedMeasurementRegistration(File mainMap, List<ModuleDocument> documents) {
        List<MeasurementCardRegistrationSheetExporter.RowData> rows = new ArrayList<>();
        for (ModuleDocument document : documents) {
            if (document.kind == ModuleKind.SOUND_INSULATION) {
                SoundInsulationProtocolDataParser.ProtocolData data =
                        SoundInsulationProtocolDataParser.parse(document.protocolFile);
                String applicationNumber = data.applicationNumber();
                if (applicationNumber.isBlank()) {
                    applicationNumber = data.registrationNumber();
                }
                rows.add(new MeasurementCardRegistrationSheetExporter.RowData(
                        applicationNumber,
                        data.registrationNumber(),
                        data.protocolNumber(),
                        "Белов Д.А."
                ));
            } else {
                rows.add(MeasurementCardRegistrationSheetExporter.rowDataFromMap(document.sourceFile, document.mapFile));
            }
        }
        MeasurementCardRegistrationSheetExporter.generateCombinedRows(
                MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mainMap), rows);
    }

    private static void generateRequestForm(ModuleDocument main,
                                            List<ModuleDocument> documents,
                                            String workDeadline,
                                            String customerInn) {
        File target = generateSingleRequestForm(main, workDeadline, customerInn);
        if (target == null || !target.exists()) {
            return;
        }
        for (int i = 1; i < documents.size(); i++) {
            TemporaryRequestForm extra = generateTemporaryRequestForm(documents.get(i), workDeadline, customerInn);
            appendDocument(target, extra.file);
            deleteTemporaryRequestForm(extra);
        }
    }

    private static File generateSingleRequestForm(ModuleDocument document, String workDeadline, String customerInn) {
        if (document.kind == ModuleKind.PHYSICAL) {
            RequestFormExporter.generate(document.sourceFile, document.mapFile, workDeadline, customerInn);
            return RequestFormExporter.resolveRequestFormFile(document.mapFile);
        }
        if (document.kind == ModuleKind.NOISE) {
            RequestFormExporter.generateForNoise(document.sourceFile, document.mapFile, workDeadline, customerInn);
            return RequestFormExporter.resolveRequestFormFile(document.mapFile);
        }
        SoundInsulationRequestFormExporter.generate(document.protocolFile, document.mapFile, workDeadline, customerInn);
        return SoundInsulationRequestFormExporter.resolveRequestFormFile(document.mapFile);
    }

    private static TemporaryRequestForm generateTemporaryRequestForm(ModuleDocument document,
                                                                     String workDeadline,
                                                                     String customerInn) {
        if (document.mapFile == null || !document.mapFile.exists()) {
            return TemporaryRequestForm.empty();
        }
        File tempDirectory = null;
        try {
            tempDirectory = Files.createTempDirectory("house-primary-request-").toFile();
            File tempMap = new File(tempDirectory, document.mapFile.getName());
            Files.copy(document.mapFile.toPath(), tempMap.toPath());
            File requestFile;
            if (document.kind == ModuleKind.PHYSICAL) {
                RequestFormExporter.generate(document.sourceFile, tempMap, workDeadline, customerInn);
                requestFile = RequestFormExporter.resolveRequestFormFile(tempMap);
            } else if (document.kind == ModuleKind.NOISE) {
                RequestFormExporter.generateForNoiseAppendix(document.sourceFile, tempMap, workDeadline, customerInn);
                requestFile = RequestFormExporter.resolveRequestFormFile(tempMap);
            } else {
                SoundInsulationMeasurementPlanExporter.generate(document.protocolFile, tempMap, workDeadline);
                SoundInsulationRequestFormExporter.generateAppendix(document.protocolFile, tempMap, workDeadline,
                        customerInn);
                requestFile = SoundInsulationRequestFormExporter.resolveRequestFormFile(tempMap);
            }
            return new TemporaryRequestForm(requestFile, tempDirectory);
        } catch (Exception ignored) {
            deleteDirectory(tempDirectory);
            return TemporaryRequestForm.empty();
        }
    }

    private static void generateRequestAnalysis(ModuleDocument main, List<ModuleDocument> documents) {
        if (main.kind != ModuleKind.SOUND_INSULATION) {
            if (main.kind == ModuleKind.NOISE) {
                RequestAnalysisSheetExporter.generateForNoise(main.mapFile, collectAdditionalMeasurementDates(documents));
            } else {
                RequestAnalysisSheetExporter.generate(main.mapFile, collectAdditionalMeasurementDates(documents));
            }
            return;
        }
        generateSingleRequestAnalysis(main);
    }

    private static List<String> collectAdditionalMeasurementDates(List<ModuleDocument> documents) {
        List<String> dates = new ArrayList<>();
        for (int i = 1; i < documents.size(); i++) {
            ModuleDocument document = documents.get(i);
            if (document.kind == ModuleKind.SOUND_INSULATION) {
                SoundInsulationProtocolDataParser.ProtocolData data =
                        SoundInsulationProtocolDataParser.parse(document.protocolFile);
                dates.addAll(data.measurementDates());
            } else {
                dates.addAll(RequestAnalysisSheetExporter.resolveMeasurementDatesForMap(document.mapFile));
            }
        }
        return dates;
    }

    private static File generateSingleRequestAnalysis(ModuleDocument document) {
        if (document.kind == ModuleKind.SOUND_INSULATION) {
            SoundInsulationRequestAnalysisSheetExporter.generate(document.protocolFile, document.mapFile);
            return SoundInsulationRequestAnalysisSheetExporter.resolveAnalysisSheetFile(document.mapFile);
        }
        if (document.kind == ModuleKind.NOISE) {
            RequestAnalysisSheetExporter.generateForNoise(document.mapFile);
            return RequestAnalysisSheetExporter.resolveAnalysisSheetFile(document.mapFile);
        }
        RequestAnalysisSheetExporter.generate(document.mapFile);
        return RequestAnalysisSheetExporter.resolveAnalysisSheetFile(document.mapFile);
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

    private static void deleteFile(File file) {
        if (file != null && file.exists()) {
            file.delete();
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
                    deleteFile(file);
                }
            }
        }
        directory.delete();
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
