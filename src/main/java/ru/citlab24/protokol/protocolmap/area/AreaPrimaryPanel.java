package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;
import ru.citlab24.protokol.protocolmap.area.noise.AreaNoisePrimaryFiles;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class AreaPrimaryPanel extends JPanel {
    private final JTextField workDeadlineField = new JTextField(18);
    private final JTextField customerInnField = new JTextField(18);
    private final JButton analyzeButton = new JButton("Начать анализ");
    private final PrimaryDropZonePanel radiationPanel;
    private final PrimaryDropZonePanel noisePanel;

    public AreaPrimaryPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать первичку по участкам");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(createTopPanel(title), BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 2, 16, 0));
        radiationPanel = new PrimaryDropZonePanel(PrimaryKind.RADIATION);
        noisePanel = new PrimaryDropZonePanel(PrimaryKind.NOISE);
        grid.add(radiationPanel);
        grid.add(noisePanel);
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
        File radiationSource = radiationPanel.hasSourceFile() ? radiationPanel.getSourceFile() : null;
        File noiseSource = noisePanel.hasSourceFile() ? noisePanel.getSourceFile() : null;
        if (radiationSource == null && noiseSource == null) {
            JOptionPane.showMessageDialog(this, "Сначала загрузите хотя бы один файл.",
                    "Нет файлов", JOptionPane.WARNING_MESSAGE);
            return;
        }

        String workDeadline = workDeadlineField.getText().trim();
        String customerInn = customerInnField.getText().trim();
        AnalysisProgressDialog progressDialog = new AnalysisProgressDialog(SwingUtilities.getWindowAncestor(this));
        analyzeButton.setEnabled(false);
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

        SwingWorker<AreaPrimaryBatchExporter.Result, Void> worker =
                new SwingWorker<AreaPrimaryBatchExporter.Result, Void>() {
            @Override
            protected AreaPrimaryBatchExporter.Result doInBackground() throws Exception {
                setProgress(8);
                AreaPrimaryBatchExporter.Result result =
                        AreaPrimaryBatchExporter.generate(radiationSource, noiseSource, workDeadline, customerInn);
                setProgress(100);
                return result;
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

    private void showBatchResults(AreaPrimaryBatchExporter.Result result) {
        if (result == null || result.documents == null) {
            return;
        }
        for (AreaPrimaryBatchExporter.ModuleDocument document : result.documents) {
            if (document.kind == AreaPrimaryBatchExporter.ModuleKind.RADIATION) {
                radiationPanel.showGeneratedPrimary(document.sourceFile, document.mapFile);
            } else {
                noisePanel.showGeneratedPrimary(document.sourceFile, document.mapFile);
            }
        }
        if (result.primaryFolder != null) {
            JOptionPane.showMessageDialog(this,
                    "Первичка сформирована в папке:\n" + result.primaryFolder.getAbsolutePath(),
                    "Готово", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    private void showBatchAnalysisError(Throwable ex) {
        JOptionPane.showMessageDialog(this,
                "Не удалось сформировать первичку:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                "Ошибка", JOptionPane.ERROR_MESSAGE);
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
            g.drawString("Собираю первичку", 42, 50);

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

    private enum PrimaryKind {
        RADIATION("Радиация", "Перетащите Excel или Word файл"),
        NOISE("Шум", "Перетащите Excel или Word файл");

        private final String title;
        private final String hint;

        PrimaryKind(String title, String hint) {
            this.title = title;
            this.hint = hint;
        }
    }

    private static class PrimaryDropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final PrimaryKind kind;
        private final JButton downloadButton;
        private final JButton resetButton;
        private File sourceFile;
        private File downloadFile;

        PrimaryDropZonePanel(PrimaryKind kind) {
            super(new BorderLayout(8, 8));
            this.kind = kind;
            setBorder(createDropBorder());
            setBackground(UIManager.getColor("Panel.background"));

            JLabel title = new JLabel(kind.title, SwingConstants.CENTER);
            title.setFont(title.getFont().deriveFont(Font.BOLD, 16f));

            JLabel hint = new JLabel(kind.hint, SwingConstants.CENTER);
            hint.setFont(hint.getFont().deriveFont(Font.PLAIN, 12f));
            hint.setForeground(UIManager.getColor("Label.disabledForeground"));

            JPanel header = new JPanel(new BorderLayout(4, 4));
            header.setOpaque(false);
            header.add(title, BorderLayout.NORTH);
            header.add(hint, BorderLayout.CENTER);

            JList<String> list = new JList<>(listModel);
            list.setVisibleRowCount(8);
            list.setBorder(BorderFactory.createEmptyBorder(6, 8, 6, 8));
            JScrollPane scrollPane = new JScrollPane(list);
            scrollPane.setBorder(BorderFactory.createEmptyBorder());

            downloadButton = new JButton("Скачать файл");
            downloadButton.setEnabled(false);
            downloadButton.addActionListener(event -> downloadGeneratedFile());

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
            this.downloadFile = null;
            listModel.clear();
            listModel.addElement("Исходный файл: " + sourceFile.getName());
            listModel.addElement("Нажмите \"Начать анализ\" для формирования документов");
            downloadButton.setEnabled(false);
            resetButton.setEnabled(true);
        }

        private void showGeneratedPrimary(File sourceFile, File mapFile) {
            this.sourceFile = sourceFile;
            listModel.clear();
            if (sourceFile != null) {
                listModel.addElement("Исходный файл: " + sourceFile.getName());
            }
            if (mapFile == null) {
                downloadButton.setEnabled(false);
                downloadFile = null;
                return;
            }
            if (kind == PrimaryKind.RADIATION) {
                showRadiationPrimary(mapFile);
            } else {
                showNoisePrimary(sourceFile, mapFile);
            }
        }

        private void showRadiationPrimary(File mapFile) {
            listModel.addElement("Сформированная карта: " + mapFile.getName());
            File registrationSheet = MeasurementCardRegistrationSheetExporter.resolveRegistrationSheetFile(mapFile);
            if (registrationSheet != null && registrationSheet.exists()) {
                listModel.addElement("Сформирован лист регистрации карт замеров: " + registrationSheet.getName());
            }
            File requestForm = RequestFormExporter.resolveRequestFormFile(mapFile);
            if (requestForm != null && requestForm.exists()) {
                listModel.addElement("Сформирована заявка: " + requestForm.getName());
            }
            File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(mapFile);
            if (analysisSheet != null && analysisSheet.exists()) {
                listModel.addElement("Сформирован лист анализа заявки: " + analysisSheet.getName());
            }
            File samplingPlan = SamplingPlanExporter.resolveSamplingPlanFile(mapFile);
            if (samplingPlan != null && samplingPlan.exists()) {
                listModel.addElement("Сформирован план отбора: " + samplingPlan.getName());
            }
            File measurementPlan = MeasurementPlanExporter.resolveMeasurementPlanFile(mapFile);
            if (measurementPlan != null && measurementPlan.exists()) {
                listModel.addElement("Сформирован план измерений: " + measurementPlan.getName());
            }
            List<File> equipmentSheets = EquipmentIssuanceSheetExporter.resolveIssuanceSheetFiles(mapFile);
            for (File equipmentSheet : equipmentSheets) {
                if (equipmentSheet != null && equipmentSheet.exists()) {
                    listModel.addElement("Сформирован лист выдачи приборов: " + equipmentSheet.getName());
                }
            }
            List<File> controlSheets = EquipmentControlSheetExporter.resolveControlSheetFiles(mapFile);
            for (File controlSheet : controlSheets) {
                if (controlSheet != null && controlSheet.exists()) {
                    listModel.addElement("Сформирован лист контроля оборудования: " + controlSheet.getName());
                }
            }
            File journalWordFile = RadiationJournalWordExporter.resolveJournalFile(mapFile);
            if (journalWordFile != null && journalWordFile.exists()) {
                listModel.addElement("Сформирован журнал (Word): " + journalWordFile.getName());
            }
            File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
            if (issuanceSheet != null && issuanceSheet.exists()) {
                listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
            }
            downloadFile = mapFile;
            downloadButton.setEnabled(downloadFile.exists());
            resetButton.setEnabled(true);
        }

        private void showNoisePrimary(File sourceFile, File mapFile) {
            listModel.addElement("Сформированная карта: " + mapFile.getName());
            File requestForm = AreaNoisePrimaryFiles.resolveRequestFormFile(mapFile);
            if (requestForm != null && requestForm.exists()) {
                listModel.addElement("Сформирована заявка: " + requestForm.getName());
            }
            File analysisSheet = AreaNoisePrimaryFiles.resolveAnalysisSheetFile(mapFile);
            if (analysisSheet != null && analysisSheet.exists()) {
                listModel.addElement("Сформирован лист анализа заявки: " + analysisSheet.getName());
            }
            File measurementPlan = AreaNoisePrimaryFiles.resolveMeasurementPlanFile(mapFile);
            if (measurementPlan != null && measurementPlan.exists()) {
                listModel.addElement("Сформирован план измерений: " + measurementPlan.getName());
            }
            File registrationSheet = AreaNoisePrimaryFiles.resolveRegistrationSheetFile(mapFile);
            if (registrationSheet != null && registrationSheet.exists()) {
                listModel.addElement("Сформирован лист регистрации карт замеров: " + registrationSheet.getName());
            }
            List<File> equipmentSheets = AreaNoisePrimaryFiles.resolveEquipmentIssuanceFiles(sourceFile, mapFile);
            for (File equipmentSheet : equipmentSheets) {
                if (equipmentSheet != null && equipmentSheet.exists()) {
                    listModel.addElement("Сформирован лист выдачи приборов: " + equipmentSheet.getName());
                }
            }
            File issuanceSheet = AreaNoisePrimaryFiles.resolveProtocolIssuanceSheetFile(mapFile);
            if (issuanceSheet != null && issuanceSheet.exists()) {
                listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
            }
            downloadFile = requestForm != null && requestForm.exists() ? requestForm : mapFile;
            downloadButton.setEnabled(downloadFile != null && downloadFile.exists());
            resetButton.setEnabled(true);
        }

        private void resetDocuments() {
            sourceFile = null;
            downloadFile = null;
            listModel.clear();
            downloadButton.setEnabled(false);
            resetButton.setEnabled(false);
        }

        private void downloadGeneratedFile() {
            if (downloadFile == null || !downloadFile.exists()) {
                JOptionPane.showMessageDialog(this, "Файл ещё не сформирован.",
                        "Нет файла", JOptionPane.WARNING_MESSAGE);
                return;
            }

            JFileChooser chooser = new JFileChooser(downloadFile.getParentFile());
            chooser.setSelectedFile(new File(downloadFile.getName()));
            int result = chooser.showSaveDialog(this);
            if (result != JFileChooser.APPROVE_OPTION) {
                return;
            }
            File target = chooser.getSelectedFile();
            try {
                Files.copy(downloadFile.toPath(), target.toPath(), StandardCopyOption.REPLACE_EXISTING);
                JOptionPane.showMessageDialog(this, "Файл сохранён: " + target.getAbsolutePath(),
                        "Готово", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "Не удалось сохранить файл: " + ex.getMessage(),
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

                    showUploadedFile(files.get(0));
                    return true;

                } catch (Exception ex) {
                    ex.printStackTrace();
                    JOptionPane.showMessageDialog(
                            PrimaryDropZonePanel.this,
                            "Ошибка при обработке файла:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                            "Ошибка",
                            JOptionPane.ERROR_MESSAGE
                    );
                    return false;
                }
            }
        }
    }
}
