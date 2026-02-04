package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.EquipmentIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementCardRegistrationSheetExporter;
import ru.citlab24.protokol.protocolmap.MeasurementPlanExporter;
import ru.citlab24.protokol.protocolmap.ProtocolIssuanceSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestAnalysisSheetExporter;
import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.List;

public class AreaPrimaryPanel extends JPanel {

    public AreaPrimaryPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать первичку по участкам");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(title, BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 2, 16, 0));
        grid.add(new PrimaryDropZonePanel(PrimaryKind.RADIATION));
        grid.add(new PrimaryDropZonePanel(PrimaryKind.NOISE));
        add(grid, BorderLayout.CENTER);
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

    private interface PrimaryGenerator {
        File generate(File sourceFile, String workDeadline, String customerInn) throws IOException;
    }

    private static class PrimaryDropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final PrimaryKind kind;
        private final PrimaryGenerator generator;
        private final JButton downloadButton;
        private File downloadFile;

        PrimaryDropZonePanel(PrimaryKind kind) {
            super(new BorderLayout(8, 8));
            this.kind = kind;
            this.generator = kind == PrimaryKind.RADIATION
                    ? AreaPrimaryRadiationExporter::generate
                    : AreaPrimaryNoiseExporter::generate;
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

        private void showGeneratedPrimary(File sourceFile, File mapFile) {
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
                showNoisePrimary(mapFile);
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
            File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
            if (issuanceSheet != null && issuanceSheet.exists()) {
                listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
            }
            downloadFile = mapFile;
            downloadButton.setEnabled(downloadFile.exists());
        }

        private void showNoisePrimary(File mapFile) {
            File shortRequest = ShortRequestFormExporter.resolveRequestFormFile(mapFile);
            if (shortRequest != null && shortRequest.exists()) {
                listModel.addElement("Сформирована заявка (краткая): " + shortRequest.getName());
            }
            File analysisSheet = RequestAnalysisSheetExporter.resolveAnalysisSheetFile(mapFile);
            if (analysisSheet != null && analysisSheet.exists()) {
                listModel.addElement("Сформирован лист анализа заявки: " + analysisSheet.getName());
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
            File issuanceSheet = ProtocolIssuanceSheetExporter.resolveIssuanceSheetFile(mapFile);
            if (issuanceSheet != null && issuanceSheet.exists()) {
                listModel.addElement("Сформирован лист выдачи протоколов: " + issuanceSheet.getName());
            }
            downloadFile = shortRequest != null && shortRequest.exists() ? shortRequest : mapFile;
            downloadButton.setEnabled(downloadFile != null && downloadFile.exists());
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

                    File source = files.get(0);
                    File generated = null;

                    String workDeadline = JOptionPane.showInputDialog(
                            PrimaryDropZonePanel.this,
                            "Какой срок выполнения работ?",
                            "Срок выполнения работ",
                            JOptionPane.QUESTION_MESSAGE
                    );
                    if (workDeadline == null) {
                        workDeadline = "";
                    }
                    String customerInn = JOptionPane.showInputDialog(
                            PrimaryDropZonePanel.this,
                            "ИНН заказчика?",
                            "ИНН заказчика",
                            JOptionPane.QUESTION_MESSAGE
                    );
                    if (customerInn == null) {
                        customerInn = "";
                    }

                    try {
                        generated = generator.generate(source, workDeadline.trim(), customerInn.trim());
                    } catch (Exception ex) {
                        ex.printStackTrace();
                        JOptionPane.showMessageDialog(
                                PrimaryDropZonePanel.this,
                                "Не удалось сформировать первичку:\n" + ex.getClass().getSimpleName() + ": " + ex.getMessage(),
                                "Ошибка",
                                JOptionPane.ERROR_MESSAGE
                        );
                    }

                    showGeneratedPrimary(source, generated);
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
