package ru.citlab24.protokol.protocolmap.house;

import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.HouseNoiseEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.NoiseMapExporter;
import ru.citlab24.protokol.protocolmap.PhysicalFactorsMapExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;
import ru.citlab24.protokol.protocolmap.SoundInsulationEquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.SoundInsulationMapExporter;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;

public class ProtocolMapPanel extends JPanel {

    public ProtocolMapPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать первичку по домам");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(title, BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 3, 16, 0));
        grid.add(new DropZonePanel("Шумы", "Перетащите Excel или Word файл",
                NoiseMapExporter::generateMap, true));
        grid.add(new DropZonePanel("Физфакторы", "Перетащите Excel или Word файл",
                PhysicalFactorsMapExporter::generateMap, false));
        grid.add(new SoundInsulationPanel());
        add(grid, BorderLayout.CENTER);
    }

    private interface MapGenerator {
        File generate(File sourceFile, String workDeadline, String customerInn) throws IOException;
    }

    private static class DropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final MapGenerator generator;
        private final JButton downloadButton;
        private final boolean isNoisePanel;
        private File generatedMapFile;

        DropZonePanel(String titleText, String hintText, MapGenerator generator, boolean isNoisePanel) {
            super(new BorderLayout(8, 8));
            this.generator = generator;
            this.isNoisePanel = isNoisePanel;
            setBorder(createDropBorder());
            setBackground(UIManager.getColor("Panel.background"));

            JLabel title = new JLabel(titleText, SwingConstants.CENTER);
            title.setFont(title.getFont().deriveFont(Font.BOLD, 16f));

            JLabel hint = new JLabel(hintText, SwingConstants.CENTER);
            hint.setFont(hint.getFont().deriveFont(Font.PLAIN, 12f));
            hint.setForeground(UIManager.getColor("Label.disabledForeground"));

            JPanel header = new JPanel(new BorderLayout(4, 4));
            header.setOpaque(false);
            header.add(title, BorderLayout.NORTH);
            header.add(hint, BorderLayout.CENTER);

            JList<String> list = new JList<>(listModel);
            list.setVisibleRowCount(4);
            list.setBorder(BorderFactory.createEmptyBorder(6, 8, 6, 8));
            JScrollPane scrollPane = new JScrollPane(list);
            scrollPane.setBorder(BorderFactory.createEmptyBorder());

            downloadButton = new JButton("Скачать карту");
            downloadButton.setEnabled(false);
            downloadButton.addActionListener(event -> downloadGeneratedMap());

            JPanel footer = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
            footer.setOpaque(false);
            footer.add(downloadButton);

            add(header, BorderLayout.NORTH);
            add(scrollPane, BorderLayout.CENTER);
            add(footer, BorderLayout.SOUTH);

            setTransferHandler(new FileDropHandler());
        }

        private static Border createDropBorder() {
            Border line = BorderFactory.createLineBorder(new Color(120, 144, 156), 1, true);
            Border padding = BorderFactory.createEmptyBorder(12, 12, 12, 12);
            return BorderFactory.createCompoundBorder(line, padding);
        }

        private void showGeneratedMap(File sourceFile, File generatedFile) {
            listModel.clear();
            listModel.addElement("Исходный файл: " + sourceFile.getName());
            String mapName = generatedFile != null ? generatedFile.getName() : buildMapName(sourceFile.getName());
            listModel.addElement("Сформированная карта: " + mapName);
            File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(generatedFile);
            if (issuanceSheet != null && issuanceSheet.exists()) {
                listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
            }
            File registrationSheet = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(generatedFile);
            if (registrationSheet != null && registrationSheet.exists()) {
                listModel.addElement("Сформирован лист регистрации карт замеров: " + registrationSheet.getName());
            }
            List<File> equipmentSheets = isNoisePanel
                    ? HouseNoiseEquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(sourceFile, generatedFile)
                    : EquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(generatedFile);
            for (File equipmentSheet : equipmentSheets) {
                if (equipmentSheet != null && equipmentSheet.exists()) {
                    listModel.addElement("Сформирован лист выдачи приборов: " + equipmentSheet.getName());
                }
            }
            File measurementPlan = MeasurementPlanExporter.resolveMeasurementPlanFile(generatedFile);
            if (measurementPlan != null && measurementPlan.exists()) {
                listModel.addElement("Сформирован план измерений: " + measurementPlan.getName());
            }
            File requestForm = RequestFormExporter.resolveRequestFormFile(generatedFile);
            if (requestForm != null && requestForm.exists()) {
                listModel.addElement("Сформирована заявка: " + requestForm.getName());
            }
            File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(generatedFile);
            if (analysisSheet != null && analysisSheet.exists()) {
                listModel.addElement("Сформирован лист анализа заявки: " + analysisSheet.getName());
            }
            generatedMapFile = generatedFile;
            downloadButton.setEnabled(generatedMapFile != null && generatedMapFile.exists());
        }

        private String buildMapName(String originalName) {
            int dotIndex = originalName.lastIndexOf('.');
            String baseName = dotIndex > 0 ? originalName.substring(0, dotIndex) : originalName;
            return baseName + "_карта.xlsx";
        }

        private void downloadGeneratedMap() {
            if (generatedMapFile == null || !generatedMapFile.exists()) {
                JOptionPane.showMessageDialog(this, "Карта ещё не сформирована.",
                        "Нет карты", JOptionPane.WARNING_MESSAGE);
                return;
            }

            JFileChooser chooser = new JFileChooser(generatedMapFile.getParentFile());
            chooser.setSelectedFile(new File(generatedMapFile.getName()));
            int result = chooser.showSaveDialog(this);
            if (result != JFileChooser.APPROVE_OPTION) {
                return;
            }
            File target = chooser.getSelectedFile();
            try {
                Files.copy(generatedMapFile.toPath(), target.toPath(), StandardCopyOption.REPLACE_EXISTING);
                JOptionPane.showMessageDialog(this, "Карта сохранена: " + target.getAbsolutePath(),
                        "Готово", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Не удалось сохранить карту: " + ex.getMessage(),
                        "Ошибка", JOptionPane.ERROR_MESSAGE);
            }
        }

        private class FileDropHandler extends TransferHandler {
            @Override
            public boolean canImport(TransferSupport support) {
                return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            public boolean importData(TransferSupport support) {
                if (!canImport(support)) {
                    return false;
                }
                try {
                    @SuppressWarnings("unchecked")
                    List<File> files = (List<File>) support.getTransferable()
                            .getTransferData(DataFlavor.javaFileListFlavor);
                    if (files == null || files.isEmpty()) {
                        return false;
                    }

                    File source = files.get(0);
                    File generated = null;

                    if (generator != null) {
                        String workDeadline = JOptionPane.showInputDialog(
                                DropZonePanel.this,
                                "Какой срок выполнения работ?",
                                "Срок выполнения работ",
                                JOptionPane.QUESTION_MESSAGE
                        );
                        if (workDeadline == null) {
                            workDeadline = "";
                        }
                        String customerInn = JOptionPane.showInputDialog(
                                DropZonePanel.this,
                                "ИНН заказчика?",
                                "ИНН заказчика",
                                JOptionPane.QUESTION_MESSAGE
                        );
                        if (customerInn == null) {
                            customerInn = "";
                        }
                        try {
                            generated = generator.generate(source, workDeadline.trim(), customerInn.trim());
                            if (isNoisePanel) {
                                HouseNoiseEquipmentIssuanceSheetExporter.generate(source, generated);
                            }
                        } catch (Exception ex) { // ВАЖНО: ловим ВСЁ, не только IOException
                            ex.printStackTrace();
                            JOptionPane.showMessageDialog(
                                    DropZonePanel.this,
                                    "Не удалось сформировать карту:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                                    "Ошибка",
                                    JOptionPane.ERROR_MESSAGE
                            );
                        }
                    }

                    showGeneratedMap(source, generated);
                    return true;

                } catch (Exception ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(
                            DropZonePanel.this,
                            "Ошибка при обработке файла:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                            "Ошибка",
                            JOptionPane.ERROR_MESSAGE
                    );
                    return false;
                }
            }
        }
    }

    private static class SoundInsulationPanel extends JPanel {
        private enum FileKind {
            IMPACT("Ударка (Excel)", "удар"),
            WALL("Стена (Excel)", "стен"),
            SLAB("Перекрытие (Excel)", "перекры"),
            PROTOCOL("Протокол (Word)", "протокол");

            private final String label;
            private final String keyword;

            FileKind(String label, String keyword) {
                this.label = label;
                this.keyword = keyword;
            }
        }

        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final Map<FileKind, List<File>> uploadedFiles = new EnumMap<>(FileKind.class);
        private final JButton analyzeButton;

        SoundInsulationPanel() {
            super(new BorderLayout(8, 8));
            setBorder(DropZonePanel.createDropBorder());
            setBackground(UIManager.getColor("Panel.background"));

            JLabel title = new JLabel("Звукоизоляция", SwingConstants.CENTER);
            title.setFont(title.getFont().deriveFont(Font.BOLD, 16f));

            JLabel hint = new JLabel("Перетащите Excel и Word файлы", SwingConstants.CENTER);
            hint.setFont(hint.getFont().deriveFont(Font.PLAIN, 12f));
            hint.setForeground(UIManager.getColor("Label.disabledForeground"));

            JPanel header = new JPanel(new BorderLayout(4, 4));
            header.setOpaque(false);
            header.add(title, BorderLayout.NORTH);
            header.add(hint, BorderLayout.CENTER);

            JList<String> list = new JList<>(listModel);
            list.setVisibleRowCount(6);
            list.setBorder(BorderFactory.createEmptyBorder(6, 8, 6, 8));
            JScrollPane scrollPane = new JScrollPane(list);
            scrollPane.setBorder(BorderFactory.createEmptyBorder());

            analyzeButton = new JButton("Начать анализ");
            analyzeButton.setEnabled(false);
            analyzeButton.addActionListener(event -> runAnalysis());

            JPanel footer = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
            footer.setOpaque(false);
            footer.add(analyzeButton);

            add(header, BorderLayout.NORTH);
            add(scrollPane, BorderLayout.CENTER);
            add(footer, BorderLayout.SOUTH);

            updateList();
            setTransferHandler(new SoundInsulationDropHandler());
        }

        private void updateList() {
            listModel.clear();
            for (FileKind kind : FileKind.values()) {
                List<File> files = uploadedFiles.get(kind);
                String fileName = "не загружен";
                if (files != null && !files.isEmpty()) {
                    fileName = files.stream()
                            .map(File::getName)
                            .collect(Collectors.joining(", "));
                }
                listModel.addElement(kind.label + ": " + fileName);
            }
        }

        private void runAnalysis() {
            if (uploadedFiles.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Сначала загрузите файлы.",
                        "Нет файлов", JOptionPane.WARNING_MESSAGE);
                return;
            }
            List<File> impactFiles = uploadedFiles.get(FileKind.IMPACT);
            List<File> wallFiles = uploadedFiles.get(FileKind.WALL);
            List<File> slabFiles = uploadedFiles.get(FileKind.SLAB);
            List<File> protocolFiles = uploadedFiles.get(FileKind.PROTOCOL);
            File protocolFile = (protocolFiles != null && !protocolFiles.isEmpty()) ? protocolFiles.get(0) : null;
            if (protocolFile == null) {
                JOptionPane.showMessageDialog(this,
                        "Не загружен протокол (Word).",
                        "Звукоизоляция",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            if (isEmpty(impactFiles) && isEmpty(wallFiles) && isEmpty(slabFiles)) {
                JOptionPane.showMessageDialog(this,
                        "Загрузите хотя бы один Excel-файл: ударка, стена или перекрытие.",
                        "Звукоизоляция",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }
            String workDeadline = JOptionPane.showInputDialog(
                    this,
                    "Какой срок выполнения работ?",
                    "Срок выполнения работ",
                    JOptionPane.QUESTION_MESSAGE
            );
            if (workDeadline == null) {
                workDeadline = "";
            }
            String customerInn = JOptionPane.showInputDialog(
                    this,
                    "ИНН заказчика?",
                    "ИНН заказчика",
                    JOptionPane.QUESTION_MESSAGE
            );
            if (customerInn == null) {
                customerInn = "";
            }
            List<File> generatedFiles = new ArrayList<>();
            List<File> impactsToProcess = impactFiles != null ? impactFiles : List.of();
            try {
                File generated = SoundInsulationMapExporter.generateMap(impactsToProcess, wallFiles, slabFiles, protocolFile,
                        workDeadline.trim(), customerInn.trim());
                if (generated != null) {
                    generatedFiles.add(generated);
                }
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(
                        this,
                        "Не удалось сформировать карту:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                        "Ошибка",
                        JOptionPane.ERROR_MESSAGE
                );
            }
            if (!generatedFiles.isEmpty()) {
                showGeneratedMaps(generatedFiles);
            }
        }

        private boolean isEmpty(List<File> files) {
            return files == null || files.isEmpty();
        }

        private void handleFiles(List<File> files) {
            boolean updated = false;
            for (File file : files) {
                FileKind kind = guessKind(file);
                if (kind == null) {
                    JOptionPane.showMessageDialog(this,
                            "Не удалось определить тип файла: " + file.getName(),
                            "Неизвестный файл",
                            JOptionPane.WARNING_MESSAGE);
                    continue;
                }
                List<File> filesByKind = uploadedFiles.computeIfAbsent(kind, key -> new ArrayList<>());
                if (!filesByKind.contains(file)) {
                    filesByKind.add(file);
                }
                updated = true;
            }
            if (updated) {
                updateList();
                analyzeButton.setEnabled(true);
            }
        }

        private FileKind guessKind(File file) {
            String rawName = file.getName();
            String name = rawName == null ? "" : rawName.trim().toLowerCase(Locale.ROOT);
            int dotIndex = name.lastIndexOf('.');
            String extension = dotIndex >= 0 ? name.substring(dotIndex + 1) : "";
            if (extension.equals("doc") || extension.equals("docx")) {
                return FileKind.PROTOCOL;
            }
            if (extension.equals("xls") || extension.equals("xlsx") || extension.equals("xlsm")) {
                if (name.contains(FileKind.IMPACT.keyword)) {
                    return FileKind.IMPACT;
                }
                if (name.contains(FileKind.WALL.keyword)
                        || name.contains("стена")
                        || name.contains("stena")
                        || name.contains("wall")) {
                    return FileKind.WALL;
                }
                if (name.contains(FileKind.SLAB.keyword)) {
                    return FileKind.SLAB;
                }
            }
            return null;
        }

        private void showGeneratedMaps(List<File> generatedFiles) {
            listModel.clear();
            for (FileKind kind : FileKind.values()) {
                List<File> files = uploadedFiles.get(kind);
                String fileName = "не загружен";
                if (files != null && !files.isEmpty()) {
                    fileName = files.stream()
                            .map(File::getName)
                            .collect(Collectors.joining(", "));
                }
                listModel.addElement(kind.label + ": " + fileName);
            }
            for (File generatedFile : generatedFiles) {
                if (generatedFile == null) {
                    continue;
                }
                listModel.addElement("Сформированная карта: " + generatedFile.getName());
                File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(generatedFile);
                if (issuanceSheet != null && issuanceSheet.exists()) {
                    listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
                }
                File registrationSheet = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(generatedFile);
                if (registrationSheet != null && registrationSheet.exists()) {
                    listModel.addElement("Сформирован лист регистрации карт замеров: " + registrationSheet.getName());
                }
                File protocolFile = null;
                List<File> protocolFiles = uploadedFiles.get(FileKind.PROTOCOL);
                if (protocolFiles != null && !protocolFiles.isEmpty()) {
                    protocolFile = protocolFiles.get(0);
                }
                List<File> equipmentSheets = protocolFile == null
                        ? EquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(generatedFile)
                        : SoundInsulationEquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(generatedFile, protocolFile);
                for (File equipmentSheet : equipmentSheets) {
                    if (equipmentSheet != null && equipmentSheet.exists()) {
                        listModel.addElement("Сформирован лист выдачи приборов: " + equipmentSheet.getName());
                    }
                }
                File measurementPlan = MeasurementPlanExporter.resolveMeasurementPlanFile(generatedFile);
                if (measurementPlan != null && measurementPlan.exists()) {
                    listModel.addElement("Сформирован план измерений: " + measurementPlan.getName());
                }
                File requestForm = RequestFormExporter.resolveRequestFormFile(generatedFile);
                if (requestForm != null && requestForm.exists()) {
                    listModel.addElement("Сформирована заявка: " + requestForm.getName());
                }
                File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(generatedFile);
                if (analysisSheet != null && analysisSheet.exists()) {
                    listModel.addElement("Сформирован лист анализа заявки: " + analysisSheet.getName());
                }
            }
        }

        private class SoundInsulationDropHandler extends TransferHandler {
            @Override
            public boolean canImport(TransferSupport support) {
                return support.isDataFlavorSupported(DataFlavor.javaFileListFlavor);
            }

            @Override
            public boolean importData(TransferSupport support) {
                if (!canImport(support)) {
                    return false;
                }
                try {
                    @SuppressWarnings("unchecked")
                    List<File> files = (List<File>) support.getTransferable()
                            .getTransferData(DataFlavor.javaFileListFlavor);
                    if (files == null || files.isEmpty()) {
                        return false;
                    }
                    handleFiles(files);
                    return true;
                } catch (Exception ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(
                            SoundInsulationPanel.this,
                            "Ошибка при обработке файлов:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                            "Ошибка",
                            JOptionPane.ERROR_MESSAGE
                    );
                    return false;
                }
            }
        }
    }
}
