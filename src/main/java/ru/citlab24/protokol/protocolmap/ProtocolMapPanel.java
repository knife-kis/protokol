package ru.citlab24.protokol.protocolmap;

import javax.swing.*;
import javax.swing.border.Border;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.io.File;
import java.util.List;

public class ProtocolMapPanel extends JPanel {

    public ProtocolMapPanel() {
        super(new BorderLayout(24, 24));
        setBorder(BorderFactory.createEmptyBorder(24, 24, 24, 24));

        JLabel title = new JLabel("Сформировать карту по протоколу");
        title.setFont(title.getFont().deriveFont(Font.BOLD, 20f));
        add(title, BorderLayout.NORTH);

        JPanel grid = new JPanel(new GridLayout(1, 3, 16, 0));
        grid.add(new DropZonePanel("Шумы", "Перетащите Excel или Word файл"));
        grid.add(new DropZonePanel("Физфакторы", "Перетащите Excel или Word файл"));
        grid.add(new DropZonePanel("Звукоизоляция", "Перетащите Excel или Word файл"));
        add(grid, BorderLayout.CENTER);
    }

    private static class DropZonePanel extends JPanel {
        private final DefaultListModel<String> listModel = new DefaultListModel<>();

        DropZonePanel(String titleText, String hintText) {
            super(new BorderLayout(8, 8));
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

            add(header, BorderLayout.NORTH);
            add(scrollPane, BorderLayout.CENTER);

            setTransferHandler(new FileDropHandler());
        }

        private static Border createDropBorder() {
            Border line = BorderFactory.createLineBorder(new Color(120, 144, 156), 1, true);
            Border padding = BorderFactory.createEmptyBorder(12, 12, 12, 12);
            return BorderFactory.createCompoundBorder(line, padding);
        }

        private void showGeneratedMap(File file) {
            listModel.clear();
            listModel.addElement("Исходный файл: " + file.getName());
            listModel.addElement("Сформированная карта: " + buildMapName(file.getName()));
        }

        private String buildMapName(String originalName) {
            int dotIndex = originalName.lastIndexOf('.');
            String baseName = dotIndex > 0 ? originalName.substring(0, dotIndex) : originalName;
            return baseName + "_карта.xlsx";
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
                    showGeneratedMap(files.get(0));
                    return true;
                } catch (Exception ex) {
                    return false;
                }
            }
        }
    }
}
