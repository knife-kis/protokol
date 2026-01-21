package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon;

import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

final class NoiseThresholdsExcelIO {

    private static final String SHEET_NAME = "Пороговые значения";
    private static final String DEFAULT_FILE = "пороги_шума_шаблон.xlsx";

    private NoiseThresholdsExcelIO() {}

    private record RowSpec(String label, String key) {}

    private static List<RowSpec> rowSpecs() {
        return List.of(
                new RowSpec("Лифт день", "Лифт|день"),
                new RowSpec("Лифт ночь", "Лифт|ночь"),
                new RowSpec("Лифт офис", "Лифт|офис"),
                new RowSpec("ИТО день", "ИТО|день"),
                new RowSpec("ИТО ночь", "ИТО|ночь"),
                new RowSpec("ИТО офис", "ИТО|офис"),
                new RowSpec("Зум день", "Зум|день"),
                new RowSpec("Авто день", "Авто|день"),
                new RowSpec("Авто ночь", "Авто|ночь"),
                new RowSpec("Улица", "Улица|диапазон")
        );
    }

    static void exportTemplate(Component parent, Map<String, double[]> thresholds) {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(SHEET_NAME);

            CellStyle header = wb.createCellStyle();
            org.apache.poi.ss.usermodel.Font font = wb.createFont();
            font.setBold(true);
            header.setFont(font);

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Источник/вариант");
            headerRow.createCell(1).setCellValue("Eq мин");
            headerRow.createCell(2).setCellValue("Eq макс");
            headerRow.createCell(3).setCellValue("MAX мин");
            headerRow.createCell(4).setCellValue("MAX макс");
            for (int i = 0; i <= 4; i++) {
                headerRow.getCell(i).setCellStyle(header);
            }

            CellStyle numberStyle = wb.createCellStyle();
            numberStyle.setDataFormat(wb.createDataFormat().getFormat("0.0"));

            int rowIndex = 1;
            for (RowSpec spec : rowSpecs()) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(spec.label());

                double[] values = (thresholds != null) ? thresholds.get(spec.key()) : null;
                writeValue(row, 1, values, 0, numberStyle);
                writeValue(row, 2, values, 1, numberStyle);
                writeValue(row, 3, values, 2, numberStyle);
                writeValue(row, 4, values, 3, numberStyle);
            }

            for (int i = 0; i <= 4; i++) {
                sheet.autoSizeColumn(i);
            }

            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить шаблон пороговых значений");
            chooser.setSelectedFile(new File(DEFAULT_FILE));
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
            if (chooser.showSaveDialog(parent) != JFileChooser.APPROVE_OPTION) return;

            File file = NoiseSheetCommon.ensureXlsx(chooser.getSelectedFile());
            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
            JOptionPane.showMessageDialog(parent, "Файл сохранён:\n" + file.getAbsolutePath(),
                    "Пороговые значения", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка сохранения: " + ex.getMessage(),
                    "Пороговые значения", JOptionPane.ERROR_MESSAGE);
        }
    }

    static Map<String, double[]> importTemplate(Component parent) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Загрузить пороговые значения");
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        if (chooser.showOpenDialog(parent) != JFileChooser.APPROVE_OPTION) {
            return null;
        }

        File file = chooser.getSelectedFile();
        Map<String, double[]> out = new LinkedHashMap<>();
        DataFormatter formatter = new DataFormatter();

        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Map<String, RowSpec> byLabel = new LinkedHashMap<>();
            for (RowSpec spec : rowSpecs()) {
                byLabel.put(spec.label().toLowerCase(), spec);
            }

            int lastRow = sheet.getLastRowNum();
            for (int i = 1; i <= lastRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String label = formatter.formatCellValue(row.getCell(0));
                if (label == null || label.trim().isEmpty()) continue;
                RowSpec spec = byLabel.get(label.trim().toLowerCase());
                if (spec == null) continue;

                double[] values = new double[4];
                values[0] = parseDoubleOrNaN(row.getCell(1), formatter);
                values[1] = parseDoubleOrNaN(row.getCell(2), formatter);
                values[2] = parseDoubleOrNaN(row.getCell(3), formatter);
                values[3] = parseDoubleOrNaN(row.getCell(4), formatter);
                if (!allNaN(values)) {
                    out.put(spec.key(), values);
                }
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка чтения файла: " + ex.getMessage(),
                    "Пороговые значения", JOptionPane.ERROR_MESSAGE);
            return null;
        }

        return out;
    }

    private static void writeValue(Row row, int cellIndex, double[] values, int arrIndex, CellStyle style) {
        if (values == null || values.length <= arrIndex) return;
        double v = values[arrIndex];
        if (Double.isNaN(v)) return;
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(v);
        cell.setCellStyle(style);
    }

    private static double parseDoubleOrNaN(Cell cell, DataFormatter formatter) {
        if (cell == null) return Double.NaN;
        String raw = formatter.formatCellValue(cell);
        if (raw == null) return Double.NaN;
        String trimmed = raw.trim().replace(',', '.');
        if (trimmed.isEmpty()) return Double.NaN;
        try {
            return Double.parseDouble(trimmed);
        } catch (NumberFormatException ex) {
            return Double.NaN;
        }
    }

    private static boolean allNaN(double[] values) {
        if (values == null) return true;
        for (double v : values) {
            if (!Double.isNaN(v)) return false;
        }
        return true;
    }
}
