package ru.citlab24.protokol.protocolmap;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.List;

public class ProtocolMapPanel extends JPanel {

    public ProtocolMapPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать карту по протоколу");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(title, BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 3, 16, 0));
        grid.add(new DropZonePanel("Шумы", "Перетащите Excel или Word файл", null));
        grid.add(new DropZonePanel("Физфакторы", "Перетащите Excel или Word файл",
                PhysicalFactorsMapExporter::generateMap));
        grid.add(new DropZonePanel("Звукоизоляция", "Перетащите Excel или Word файл", null));
        add(grid, BorderLayout.CENTER);
    }

    private interface MapGenerator {
        File generate(File sourceFile) throws IOException;
    }

    private static class DropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();
        private final MapGenerator generator;
        private final JButton downloadButton;
        private File generatedMapFile;

        DropZonePanel(String titleText, String hintText, MapGenerator generator) {
            super(new BorderLayout(8, 8));
            this.generator = generator;
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
                        try {
                            generated = generator.generate(source);
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
}
