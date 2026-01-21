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
import java.time.LocalDate;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.EnumMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

final class NoisePeriodsExcelIO {

    private static final String SHEET_NAME = "Периоды измерений";
    private static final String DEFAULT_FILE = "периоды_шума.xlsx";
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final DateTimeFormatter TIME_FORMAT = DateTimeFormatter.ofPattern("HH:mm");

    private NoisePeriodsExcelIO() {}

    private record RowSpec(String label, NoiseTestKind kind) {}

    private static List<RowSpec> rowSpecs() {
        return List.of(
                new RowSpec("Лифт — день", NoiseTestKind.LIFT_DAY),
                new RowSpec("Лифт — ночь", NoiseTestKind.LIFT_NIGHT),
                new RowSpec("ИТО — нежилые", NoiseTestKind.ITO_NONRES),
                new RowSpec("ИТО — жилые день", NoiseTestKind.ITO_RES_DAY),
                new RowSpec("ИТО — жилые ночь", NoiseTestKind.ITO_RES_NIGHT),
                new RowSpec("Авто — день", NoiseTestKind.AUTO_DAY),
                new RowSpec("Авто — ночь", NoiseTestKind.AUTO_NIGHT),
                new RowSpec("Площадка (улица)", NoiseTestKind.SITE)
        );
    }

    static void exportPeriods(Component parent,
                              Map<NoiseTestKind, NoisePeriod> periods,
                              Map<NoiseTestKind, Integer> pointsByKind) {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet(SHEET_NAME);

            CellStyle header = wb.createCellStyle();
            org.apache.poi.ss.usermodel.Font font = wb.createFont();
            font.setBold(true);
            header.setFont(font);

            Row headerRow = sheet.createRow(0);
            headerRow.createCell(0).setCellValue("Вид измерения");
            headerRow.createCell(1).setCellValue("Дата");
            headerRow.createCell(2).setCellValue("Время с");
            headerRow.createCell(3).setCellValue("Время до");
            headerRow.createCell(4).setCellValue("Что измеряем");
            for (int i = 0; i <= 4; i++) {
                headerRow.getCell(i).setCellStyle(header);
            }

            CellStyle dateStyle = wb.createCellStyle();
            dateStyle.setDataFormat(wb.createDataFormat().getFormat("dd.mm.yyyy"));
            CellStyle timeStyle = wb.createCellStyle();
            timeStyle.setDataFormat(wb.createDataFormat().getFormat("hh:mm"));

            int rowIndex = 1;
            for (RowSpec spec : rowSpecs()) {
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(spec.label());

                NoisePeriod period = (periods != null) ? periods.get(spec.kind()) : null;
                if (period != null) {
                    writeDate(row, 1, period.date, dateStyle);
                    writeTime(row, 2, period.from, timeStyle);
                    writeTime(row, 3, period.to, timeStyle);
                }

                String description = measurementDescription(pointsByKind, spec.kind());
                if (!description.isEmpty()) {
                    row.createCell(4).setCellValue(description);
                }
            }

            for (int i = 0; i <= 4; i++) {
                sheet.autoSizeColumn(i);
            }

            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить периоды измерений");
            chooser.setSelectedFile(new File(DEFAULT_FILE));
            chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
            if (chooser.showSaveDialog(parent) != JFileChooser.APPROVE_OPTION) return;

            File file = NoiseSheetCommon.ensureXlsx(chooser.getSelectedFile());
            try (FileOutputStream out = new FileOutputStream(file)) {
                wb.write(out);
            }
            JOptionPane.showMessageDialog(parent, "Файл сохранён:\n" + file.getAbsolutePath(),
                    "Периоды измерений", JOptionPane.INFORMATION_MESSAGE);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка сохранения: " + ex.getMessage(),
                    "Периоды измерений", JOptionPane.ERROR_MESSAGE);
        }
    }

    static Map<NoiseTestKind, NoisePeriod> importPeriods(Component parent) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Загрузить периоды измерений");
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        if (chooser.showOpenDialog(parent) != JFileChooser.APPROVE_OPTION) {
            return null;
        }

        File file = chooser.getSelectedFile();
        Map<NoiseTestKind, NoisePeriod> out = new EnumMap<>(NoiseTestKind.class);
        DataFormatter formatter = new DataFormatter();

        try (Workbook wb = WorkbookFactory.create(file)) {
            Sheet sheet = wb.getSheetAt(0);
            Map<String, NoiseTestKind> byLabel = new LinkedHashMap<>();
            for (RowSpec spec : rowSpecs()) {
                byLabel.put(spec.label().toLowerCase(), spec.kind());
            }
            byLabel.put("зум — день", NoiseTestKind.ITO_RES_DAY);
            byLabel.put("зум - день", NoiseTestKind.ITO_RES_DAY);

            int lastRow = sheet.getLastRowNum();
            for (int i = 1; i <= lastRow; i++) {
                Row row = sheet.getRow(i);
                if (row == null) continue;
                String label = formatter.formatCellValue(row.getCell(0));
                if (label == null || label.trim().isEmpty()) continue;
                NoiseTestKind kind = byLabel.get(label.trim().toLowerCase());
                if (kind == null) continue;

                LocalDate date = parseDate(row.getCell(1), formatter);
                LocalTime from = parseTime(row.getCell(2), formatter);
                LocalTime to = parseTime(row.getCell(3), formatter);
                if (date != null || from != null || to != null) {
                    out.put(kind, new NoisePeriod(date, from, to));
                }
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка чтения файла: " + ex.getMessage(),
                    "Периоды измерений", JOptionPane.ERROR_MESSAGE);
            return null;
        }

        return out;
    }

    private static String measurementDescription(Map<NoiseTestKind, Integer> pointsByKind, NoiseTestKind kind) {
        if (kind == null) return "";
        String label = kindLabel(kind);
        StringBuilder sb = new StringBuilder(label);
        if (pointsByKind != null) {
            Integer points = pointsByKind.get(kind);
            if (points != null && points > 0) {
                sb.append(" (точек: ").append(points).append(")");
            }
        }
        return sb.toString();
    }

    private static String kindLabel(NoiseTestKind kind) {
        return switch (kind) {
            case LIFT_DAY -> "Шум лифта — день";
            case LIFT_NIGHT -> "Шум лифта — ночь";
            case ITO_NONRES -> "Шум ИТО — нежилые";
            case ITO_RES_DAY -> "Шум ИТО — жилые день";
            case ITO_RES_NIGHT -> "Шум ИТО — жилые ночь";
            case ZUM_DAY -> "Шум ИТО — жилые день";
            case AUTO_DAY -> "Шум авто — день";
            case AUTO_NIGHT -> "Шум авто — ночь";
            case SITE -> "Шум площадка (улица)";
        };
    }

    private static void writeDate(Row row, int cellIndex, LocalDate date, CellStyle style) {
        if (date == null) return;
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(date.format(DATE_FORMAT));
        cell.setCellStyle(style);
    }

    private static void writeTime(Row row, int cellIndex, LocalTime time, CellStyle style) {
        if (time == null) return;
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(time.format(TIME_FORMAT));
        cell.setCellStyle(style);
    }

    private static LocalDate parseDate(Cell cell, DataFormatter formatter) {
        String raw = formatter.formatCellValue(cell);
        if (raw == null) return null;
        String trimmed = raw.trim();
        if (trimmed.isEmpty()) return null;
        try {
            return LocalDate.parse(trimmed, DATE_FORMAT);
        } catch (Exception ex) {
            return null;
        }
    }

    private static LocalTime parseTime(Cell cell, DataFormatter formatter) {
        String raw = formatter.formatCellValue(cell);
        if (raw == null) return null;
        String trimmed = raw.trim();
        if (trimmed.isEmpty()) return null;
        try {
            return LocalTime.parse(trimmed, TIME_FORMAT);
        } catch (Exception ex) {
            return null;
        }
    }
}
