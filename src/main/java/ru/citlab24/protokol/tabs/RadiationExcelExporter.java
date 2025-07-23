package ru.citlab24.protokol.tabs;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class RadiationExcelExporter {

    public static void export(List<RadiationRecord> records, java.awt.Component parent) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("МЭД");

        // Базовый шрифт Arial 10pt
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        // Стили
        CellStyle headerStyle = createHeaderStyle(workbook, baseFont);
        CellStyle titleStyle = createTitleStyle(workbook, baseFont);
        CellStyle dataStyle = createDataStyle(workbook, baseFont);
        CellStyle floorHeaderStyle = createFloorHeaderStyle(workbook, baseFont);
        CellStyle numberStyle = createNumberStyle(workbook, baseFont, "0.00");

        // Создаем структуру документа
        createDocumentStructure(sheet, titleStyle);
        createTableHeaders(sheet, headerStyle);
        createColumnNumbers(sheet, headerStyle);

        // Устанавливаем ширину колонок
        setColumnsWidth(sheet);

        // Устанавливаем высоту строк
        setRowsHeight(sheet);

        fillData(records, sheet, dataStyle, numberStyle, floorHeaderStyle);

        // Сохранение файла
        saveWorkbook(workbook, parent);
    }

    private static void setColumnsWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 31 * 256 / 7); // № п/п
        sheet.setColumnWidth(1, 150 * 256 / 7); // Помещение
        sheet.setColumnWidth(2, 100 * 256 / 7); // Результат измерения
        sheet.setColumnWidth(3, 100 * 256 / 7); // Допустимый уровень
    }

    private static void setRowsHeight(Sheet sheet) {
        if (sheet.getRow(0) != null) sheet.getRow(0).setHeightInPoints(20);
        if (sheet.getRow(1) != null) sheet.getRow(1).setHeightInPoints(15);
        if (sheet.getRow(2) != null) sheet.getRow(2).setHeightInPoints(20);
        if (sheet.getRow(3) != null) sheet.getRow(3).setHeightInPoints(15);
    }

    private static CellStyle createHeaderStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        style.setWrapText(true);
        setBlackBorderColor(style);
        return style;
    }

    private static CellStyle createTitleStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(baseFont);
        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        setBlackBorderColor(style);
        return style;
    }

    private static CellStyle createNumberStyle(Workbook workbook, Font baseFont, String format) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        setBlackBorderColor(style);
        return style;
    }

    private static CellStyle createFloorHeaderStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(baseFont);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        setBlackBorderColor(style);
        return style;
    }

    private static void setBlackBorderColor(CellStyle style) {
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
    }

    private static void createDocumentStructure(Sheet sheet, CellStyle titleStyle) {
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("15.5. Мощность амбиентного эквивалента дозы гамма-излучения на рабочих местах");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        sheet.createRow(1);
    }

    private static void createTableHeaders(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(2);
        String[] headers = {
                "№ п/п", "Помещение", "Результат измерения, мкЗв/ч", "Допустимый уровень, мкЗв/ч"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    private static void createColumnNumbers(Sheet sheet, CellStyle headerStyle) {
        Row numberRow = sheet.createRow(3);
        for (int i = 0; i < 4; i++) {
            Cell cell = numberRow.createCell(i);
            cell.setCellValue(i + 1);
            cell.setCellStyle(headerStyle);
        }
    }

    private static void fillData(List<RadiationRecord> records, Sheet sheet,
                                 CellStyle dataStyle, CellStyle numberStyle,
                                 CellStyle floorHeaderStyle) {
        int rowNum = 4;
        int counter = 1;

        // Группировка записей по этажам
        Map<String, List<RadiationRecord>> recordsByFloor = records.stream()
                .collect(Collectors.groupingBy(
                        r -> r.getBuilding() + "|" + r.getFloor(), // Уникальный ключ
                        Collectors.toList()
                ));

        for (Map.Entry<String, List<RadiationRecord>> entry : recordsByFloor.entrySet()) {
            String floor = entry.getKey();
            List<RadiationRecord> floorRecords = entry.getValue();

            // Заголовок этажа
            Row floorRow = sheet.createRow(rowNum++);
            floorRow.setHeightInPoints(15);
            Cell floorCell = floorRow.createCell(0);
            floorCell.setCellValue(floor);
            floorCell.setCellStyle(floorHeaderStyle);
            CellRangeAddress mergedRegion = new CellRangeAddress(rowNum-1, rowNum-1, 0, 3);
            sheet.addMergedRegion(mergedRegion);
            RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegion, sheet);

            for (RadiationRecord record : floorRecords) {
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.setHeightInPoints(15);

                // Колонка A
                dataRow.createCell(0).setCellValue(counter++);
                dataRow.getCell(0).setCellStyle(dataStyle);

                // Колонка B
                dataRow.createCell(1).setCellValue(record.getRoom());
                dataRow.getCell(1).setCellStyle(dataStyle);

                // Колонка C
                Cell resultCell = dataRow.createCell(2);
                if (record.getMeasurementResult() != null) {
                    resultCell.setCellValue(record.getMeasurementResult());
                } else {
                    resultCell.setCellValue("-");
                }
                resultCell.setCellStyle(numberStyle);

                // Колонка D
                Cell levelCell = dataRow.createCell(3);
                if (record.getPermissibleLevel() != null) {
                    levelCell.setCellValue(record.getPermissibleLevel());
                } else {
                    levelCell.setCellValue("-");
                }
                levelCell.setCellStyle(numberStyle);
            }
        }
    }

    private static void saveWorkbook(Workbook workbook, java.awt.Component parent) {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Сохранить отчет");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Excel Files", "xlsx"));

        if (fileChooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                file = new File(file.getAbsolutePath() + ".xlsx");
            }

            try (FileOutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
                JOptionPane.showMessageDialog(parent,
                        "Файл успешно сохранен: " + file.getAbsolutePath(),
                        "Экспорт завершен",
                        JOptionPane.INFORMATION_MESSAGE);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(parent,
                        "Ошибка при сохранении файла: " + e.getMessage(),
                        "Ошибка",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }
}
