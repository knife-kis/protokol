package ru.citlab24.protokol.protocolmap.house;

import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.HousePrimaryBatchExporter;
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
import java.util.concurrent.ExecutionException;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.stream.Collectors;

public class ProtocolMapPanel extends JPanel {
    private static final String HOUSE_PRIMARY_FOLDER_NAME = "первичка дома";
    private final JTextField workDeadlineField = new JTextField(18);
    private final JTextField customerInnField = new JTextField(18);
    private final JButton analyzeButton = new JButton("Начать анализ");
    private final DropZonePanel noisePanel;
    private final DropZonePanel physicalPanel;
    private final SoundInsulationPanel soundInsulationPanel;

    public ProtocolMapPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать первичку по домам");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(createTopPanel(title), BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 3, 16, 0));
        noisePanel = new DropZonePanel("Шумы", "Перетащите Excel или Word файл", true);
        physicalPanel = new DropZonePanel("Физфакторы", "Перетащите Excel или Word файл", false);
        soundInsulationPanel = new SoundInsulationPanel();
        grid.add(noisePanel);
        grid.add(physicalPanel);
        grid.add(soundInsulationPanel);
        add(grid, BorderLayout.CENTER);
    }

    private JPanel createTopPanel(JLabel title) {
        JPanel panel = new JPanel(new BorderLayout(12, 8));
        JPanel controls = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        controls.add(new JLabel("Срок выполнения работ:"));
        controls.add(workDeadlineField);
        controls.add(new JLabel("ИНН заказчика:"));
        controls.add(customerInnField);
        analyzeButton.addActionListener(event -> runBatchAnalysis());
        controls.add(analyzeButton);
        panel.add(title, BorderLayout.NORTH);
        panel.add(controls, BorderLayout.SOUTH);
        return panel;
    }

    private void runBatchAnalysis() {
        String workDeadline = workDeadlineField.getText().trim();
        String customerInn = customerInnField.getText().trim();
        File physicalSource = physicalPanel.hasSourceFile() ? physicalPanel.getSourceFile() : null;
        File noiseSource = noisePanel.hasSourceFile() ? noisePanel.getSourceFile() : null;
        boolean hasSoundInsulation = soundInsulationPanel.hasDocuments();
        if (physicalSource == null && noiseSource == null && !hasSoundInsulation) {
            JOptionPane.showMessageDialog(this, "Сначала загрузите хотя бы один файл.",
                    "Нет файлов", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (hasSoundInsulation && !soundInsulationPanel.validateReady(this)) {
            return;
        }

        List<File> impactFiles = new ArrayList<>(soundInsulationPanel.getImpactFiles());
        List<File> wallFiles = new ArrayList<>(soundInsulationPanel.getWallFiles());
        List<File> slabFiles = new ArrayList<>(soundInsulationPanel.getSlabFiles());
        File soundProtocolFile = soundInsulationPanel.getProtocolFile();
        File soundSource = soundInsulationPanel.getSourceFile();

        AnalysisProgressDialog progressDialog = new AnalysisProgressDialog(SwingUtilities.getWindowAncestor(this));
        analyzeButton.setEnabled(false);
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        SwingWorker<List<HousePrimaryBatchExporter.ModuleDocument>, Void> worker =
                new SwingWorker<List<HousePrimaryBatchExporter.ModuleDocument>, Void>() {
            @Override
            protected List<HousePrimaryBatchExporter.ModuleDocument> doInBackground() throws Exception {
                return createBatchDocuments(workDeadline, customerInn, physicalSource, noiseSource,
                        impactFiles, wallFiles, slabFiles, soundProtocolFile, soundSource,
                        progress -> setProgress(progress));
            }

            @Override
            protected void done() {
                progressDialog.stopAnimation();
                analyzeButton.setEnabled(true);
                setCursor(Cursor.getDefaultCursor());
                try {
                    showBatchResults(get());
                } catch (InterruptedException ex) {
                    Thread.currentThread().interrupt();
                    showBatchAnalysisError(ex);
                } catch (ExecutionException ex) {
                    Throwable cause = ex.getCause() == null ? ex : ex.getCause();
                    showBatchAnalysisError(cause);
                }
            }
        };
        worker.addPropertyChangeListener(event -> {
            if ("progress".equals(event.getPropertyName())) {
                progressDialog.setProgress((Integer) event.getNewValue());
            }
        });
        worker.execute();
        progressDialog.startAnimation();
    }

    private List<HousePrimaryBatchExporter.ModuleDocument> createBatchDocuments(String workDeadline,
                                                                               String customerInn,
                                                                               File physicalSource,
                                                                               File noiseSource,
                                                                               List<File> impactFiles,
                                                                               List<File> wallFiles,
                                                                               List<File> slabFiles,
                                                                               File soundProtocolFile,
                                                                               File soundSource,
                                                                               ProgressReporter progressReporter) throws IOException {
        List<HousePrimaryBatchExporter.ModuleDocument> documents = new ArrayList<>();
        int totalStages = (physicalSource != null ? 1 : 0)
                + (noiseSource != null ? 1 : 0)
                + (soundSource != null ? 1 : 0)
                + 1;
        int completedStages = 0;
        reportProgress(progressReporter, 3);
        File primaryFolder = null;
        if (physicalSource != null) {
            primaryFolder = ensureHousePrimaryFolder(primaryFolder, physicalSource);
            File generated = PhysicalFactorsMapExporter.generateMap(physicalSource, workDeadline, customerInn,
                    primaryFolder, false);
            documents.add(new HousePrimaryBatchExporter.ModuleDocument(
                    HousePrimaryBatchExporter.ModuleKind.PHYSICAL, physicalSource, generated, null));
            reportProgress(progressReporter, ++completedStages * 100 / totalStages);
        }
        if (noiseSource != null) {
            primaryFolder = ensureHousePrimaryFolder(primaryFolder, noiseSource);
            File generated = NoiseMapExporter.generateMap(noiseSource, workDeadline, customerInn,
                    primaryFolder, false);
            documents.add(new HousePrimaryBatchExporter.ModuleDocument(
                    HousePrimaryBatchExporter.ModuleKind.NOISE, noiseSource, generated, null));
            reportProgress(progressReporter, ++completedStages * 100 / totalStages);
        }
        if (soundSource != null) {
            primaryFolder = ensureHousePrimaryFolder(primaryFolder, soundSource);
            File generated = SoundInsulationMapExporter.generateMap(
                    impactFiles,
                    wallFiles,
                    slabFiles,
                    soundProtocolFile,
                    workDeadline,
                    customerInn,
                    primaryFolder,
                    false);
            documents.add(new HousePrimaryBatchExporter.ModuleDocument(
                    HousePrimaryBatchExporter.ModuleKind.SOUND_INSULATION,
                    soundSource,
                    generated,
                    soundProtocolFile));
            reportProgress(progressReporter, ++completedStages * 100 / totalStages);
        }
        HousePrimaryBatchExporter.generateCompanionDocuments(documents, workDeadline, customerInn);
        reportProgress(progressReporter, 100);
        return documents;
    }

    private static void reportProgress(ProgressReporter progressReporter, int progress) {
        if (progressReporter != null) {
            progressReporter.report(Math.max(0, Math.min(100, progress)));
        }
    }

    private interface ProgressReporter {
        void report(int progress);
    }

    private void showBatchAnalysisError(Throwable ex) {
        ex.printStackTrace();
        JOptionPane.showMessageDialog(this,
                "Не удалось выполнить анализ:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                "Ошибка",
                JOptionPane.ERROR_MESSAGE);
    }

    private File ensureHousePrimaryFolder(File currentFolder, File sourceFile) {
        if (currentFolder != null) {
            return currentFolder;
        }
        File parent = sourceFile != null ? sourceFile.getParentFile() : null;
        if (parent == null) {
            parent = new File(".");
        }
        File folder = new File(parent, HOUSE_PRIMARY_FOLDER_NAME);
        if (!folder.exists()) {
            folder.mkdirs();
        }
        return folder;
    }

    private void showBatchResults(List<HousePrimaryBatchExporter.ModuleDocument> documents) {
        for (HousePrimaryBatchExporter.ModuleDocument document : documents) {
            if (document.kind == HousePrimaryBatchExporter.ModuleKind.PHYSICAL) {
                physicalPanel.showGeneratedMap(document.sourceFile, document.mapFile);
            } else if (document.kind == HousePrimaryBatchExporter.ModuleKind.NOISE) {
                noisePanel.showGeneratedMap(document.sourceFile, document.mapFile);
            } else if (document.kind == HousePrimaryBatchExporter.ModuleKind.SOUND_INSULATION) {
                soundInsulationPanel.showGeneratedMaps(List.of(document.mapFile));
            }
        }
    }

    private static class AnalysisProgressDialog extends JDialog {
        private final ConstructionLoadingPanel loadingPanel = new ConstructionLoadingPanel();

        AnalysisProgressDialog(Window owner) {
            super(owner, "Анализ", Dialog.ModalityType.MODELESS);
            setDefaultCloseOperation(WindowConstants.DO_NOTHING_ON_CLOSE);
            setResizable(false);
            setContentPane(loadingPanel);
            pack();
            setLocationRelativeTo(owner);
        }

        private void startAnimation() {
            loadingPanel.start();
            setVisible(true);
        }

        private void setProgress(int progress) {
            loadingPanel.setProgress(progress);
        }

        private void stopAnimation() {
            loadingPanel.stop();
            setVisible(false);
            dispose();
        }
    }

    private static class ConstructionLoadingPanel extends JPanel {
        private final javax.swing.Timer timer;
        private int frame;
        private int progress;
        private long startedAtMillis;
        private long estimateUpdatedAtMillis;
        private long remainingMillisEstimate;

        ConstructionLoadingPanel() {
            setPreferredSize(new Dimension(430, 210));
            setBackground(new Color(246, 249, 252));
            setBorder(BorderFactory.createEmptyBorder(18, 22, 18, 22));
            timer = new javax.swing.Timer(45, event -> {
                frame++;
                repaint();
            });
        }

        private void start() {
            frame = 0;
            progress = 0;
            startedAtMillis = System.currentTimeMillis();
            estimateUpdatedAtMillis = startedAtMillis;
            remainingMillisEstimate = -1L;
            timer.start();
        }

        private void setProgress(int progress) {
            int newProgress = Math.max(0, Math.min(100, progress));
            long now = System.currentTimeMillis();
            if (newProgress > this.progress && newProgress > 3 && newProgress < 100 && startedAtMillis > 0) {
                long elapsed = Math.max(1L, now - startedAtMillis);
                long estimatedTotal = elapsed * 100L / newProgress;
                remainingMillisEstimate = Math.max(0L, estimatedTotal - elapsed);
                estimateUpdatedAtMillis = now;
            }
            this.progress = newProgress;
            repaint();
        }

        private void stop() {
            timer.stop();
        }

        @Override
        protected void paintComponent(Graphics graphics) {
            super.paintComponent(graphics);
            Graphics2D g = (Graphics2D) graphics.create();
            g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

            int width = getWidth();
            int height = getHeight();
            paintBlueprintGrid(g, width, height);
            paintBlocks(g);
            paintProgress(g, width);
            paintText(g, width);

            g.dispose();
        }

        private void paintBlueprintGrid(Graphics2D g, int width, int height) {
            g.setColor(new Color(226, 235, 244));
            for (int x = 18; x < width; x += 24) {
                g.drawLine(x, 0, x, height);
            }
            for (int y = 16; y < height; y += 24) {
                g.drawLine(0, y, width, y);
            }
            g.setColor(new Color(210, 225, 238));
            g.drawRoundRect(18, 16, width - 36, height - 32, 18, 18);
        }

        private void paintBlocks(Graphics2D g) {
            Color[] colors = {
                    new Color(44, 139, 220),
                    new Color(255, 184, 76),
                    new Color(96, 181, 122),
                    new Color(232, 94, 84),
                    new Color(134, 117, 217)
            };
            int baseX = 58;
            int baseY = 128;
            int size = 24;
            int gap = 7;
            int cycle = frame % 132;
            for (int i = 0; i < 10; i++) {
                int column = i % 5;
                int level = i / 5;
                double phase = clamp((cycle - i * 8) / 34.0, 0.0, 1.0);
                double eased = 1 - Math.pow(1 - phase, 3);
                int targetX = baseX + column * (size + gap);
                int targetY = baseY - level * (size + gap);
                int sourceX = baseX + 264 + (i % 2) * 18;
                int sourceY = baseY + 16 - i * 4;
                int x = (int) Math.round(sourceX + (targetX - sourceX) * eased);
                int y = (int) Math.round(sourceY + (targetY - sourceY) * eased);
                int alpha = (int) Math.round(80 + 175 * eased);
                g.setColor(new Color(0, 0, 0, 28));
                g.fillRoundRect(x + 3, y + 5, size, size, 7, 7);
                Color color = colors[i % colors.length];
                g.setColor(new Color(color.getRed(), color.getGreen(), color.getBlue(), alpha));
                g.fillRoundRect(x, y, size, size, 7, 7);
                g.setColor(new Color(255, 255, 255, 95));
                g.drawLine(x + 6, y + 6, x + size - 7, y + 6);
                g.setColor(new Color(48, 58, 72, 130));
                g.drawRoundRect(x, y, size, size, 7, 7);
            }

            g.setStroke(new BasicStroke(3f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND));
            g.setColor(new Color(64, 83, 105, 120));
            g.drawLine(baseX - 8, baseY + size + 8, baseX + 5 * (size + gap), baseY + size + 8);
        }

        private void paintProgress(Graphics2D g, int width) {
            int x = 42;
            int y = 165;
            int barWidth = width - 84;
            int barHeight = 12;
            g.setColor(new Color(218, 230, 240));
            g.fillRoundRect(x, y, barWidth, barHeight, 10, 10);
            g.setColor(new Color(44, 139, 220));
            g.fillRoundRect(x, y, Math.max(8, barWidth * progress / 100), barHeight, 10, 10);
            g.setColor(new Color(64, 83, 105, 90));
            g.drawRoundRect(x, y, barWidth, barHeight, 10, 10);
        }

        private void paintText(Graphics2D g, int width) {
            g.setColor(new Color(39, 51, 66));
            g.setFont(getFont().deriveFont(Font.BOLD, 18f));
            String title = "Собираю первичку";
            g.drawString(title, 42, 50);

            g.setFont(getFont().deriveFont(Font.PLAIN, 12f));
            g.setColor(new Color(85, 102, 121));
            g.drawString("Кирпичики документов укладываются по местам", 42, 73);

            int dotCount = (frame / 10) % 4;
            String dots = dotCount == 0 ? "" : dotCount == 1 ? "." : dotCount == 2 ? ".." : "...";
            String status = "Идет анализ" + dots + "  " + progress + "%";
            String remaining = buildRemainingText();
            FontMetrics metrics = g.getFontMetrics();
            g.setColor(new Color(44, 139, 220));
            g.drawString(status, 42, 192);
            g.setColor(new Color(85, 102, 121));
            g.drawString(remaining, width - metrics.stringWidth(remaining) - 42, 192);
        }

        private String buildRemainingText() {
            if (progress <= 3 || startedAtMillis <= 0) {
                return "оцениваю время";
            }
            if (progress >= 100) {
                return "завершаю";
            }
            if (remainingMillisEstimate < 0L) {
                return "оцениваю время";
            }
            long passedAfterEstimate = Math.max(0L, System.currentTimeMillis() - estimateUpdatedAtMillis);
            long remaining = Math.max(0L, remainingMillisEstimate - passedAfterEstimate);
            long seconds = Math.max(1L, Math.round(remaining / 1000.0));
            if (seconds < 60) {
                return "осталось примерно " + seconds + " с";
            }
            return "осталось примерно " + (seconds / 60) + " мин " + (seconds % 60) + " с";
        }

        private double clamp(double value, double min, double max) {
            return Math.max(min, Math.min(max, value));
        }
    }

    private static class DropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final JButton downloadButton;
        private final JButton resetButton;
        private final boolean isNoisePanel;
        private File sourceFile;
        private File generatedMapFile;

        DropZonePanel(String titleText, String hintText, boolean isNoisePanel) {
            super(new BorderLayout(8, 8));
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

            resetButton = new JButton("Сбросить документы");
            resetButton.setEnabled(false);
            resetButton.addActionListener(event -> resetDocuments());

            JPanel footer = new JPanel(new FlowLayout(FlowLayout.CENTER, 8, 0));
            footer.setOpaque(false);
            footer.add(downloadButton);
            footer.add(resetButton);

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

        private boolean hasSourceFile() {
            return sourceFile != null && sourceFile.exists();
        }

        private File getSourceFile() {
            return sourceFile;
        }

        private void showUploadedFile(File sourceFile) {
            this.sourceFile = sourceFile;
            this.generatedMapFile = null;
            listModel.clear();
            listModel.addElement("Исходный файл: " + sourceFile.getName());
            listModel.addElement("Нажмите \"Начать анализ\" для формирования документов");
            downloadButton.setEnabled(false);
            resetButton.setEnabled(true);
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
            File measurementPlan = isNoisePanel
                    ? MeasurementPlanExporter.resolveNoiseMeasurementPlanFile(generatedFile)
                    : MeasurementPlanExporter.resolveMeasurementPlanFile(generatedFile);
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
            resetButton.setEnabled(true);
        }

        private void resetDocuments() {
            sourceFile = null;
            generatedMapFile = null;
            listModel.clear();
            downloadButton.setEnabled(false);
            resetButton.setEnabled(false);
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
                    showUploadedFile(source);
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
        private final JButton resetButton;

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

            resetButton = new JButton("Сбросить документы");
            resetButton.setEnabled(false);
            resetButton.addActionListener(event -> resetDocuments());

            JPanel footer = new JPanel(new FlowLayout(FlowLayout.CENTER, 8, 0));
            footer.setOpaque(false);
            footer.add(resetButton);

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

        private void resetDocuments() {
            uploadedFiles.clear();
            updateList();
            resetButton.setEnabled(false);
        }

        private boolean isEmpty(List<File> files) {
            return files == null || files.isEmpty();
        }

        private boolean hasDocuments() {
            return !uploadedFiles.isEmpty();
        }

        private boolean validateReady(Component parent) {
            if (getProtocolFile() == null) {
                JOptionPane.showMessageDialog(parent,
                        "Не загружен протокол (Word) для звукоизоляции.",
                        "Звукоизоляция",
                        JOptionPane.WARNING_MESSAGE);
                return false;
            }
            if (isEmpty(getImpactFiles()) && isEmpty(getWallFiles()) && isEmpty(getSlabFiles())) {
                JOptionPane.showMessageDialog(parent,
                        "Загрузите хотя бы один Excel-файл звукоизоляции: ударка, стена или перекрытие.",
                        "Звукоизоляция",
                        JOptionPane.WARNING_MESSAGE);
                return false;
            }
            return true;
        }

        private List<File> getImpactFiles() {
            return uploadedFiles.getOrDefault(FileKind.IMPACT, List.of());
        }

        private List<File> getWallFiles() {
            return uploadedFiles.getOrDefault(FileKind.WALL, List.of());
        }

        private List<File> getSlabFiles() {
            return uploadedFiles.getOrDefault(FileKind.SLAB, List.of());
        }

        private File getProtocolFile() {
            List<File> protocolFiles = uploadedFiles.get(FileKind.PROTOCOL);
            return protocolFiles == null || protocolFiles.isEmpty() ? null : protocolFiles.get(0);
        }

        private File getSourceFile() {
            if (!isEmpty(getImpactFiles())) {
                return getImpactFiles().get(0);
            }
            if (!isEmpty(getWallFiles())) {
                return getWallFiles().get(0);
            }
            if (!isEmpty(getSlabFiles())) {
                return getSlabFiles().get(0);
            }
            return null;
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
                resetButton.setEnabled(true);
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
                File measurementPlan = SoundInsulationMapExporter.resolveMeasurementPlanFile(generatedFile);
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
