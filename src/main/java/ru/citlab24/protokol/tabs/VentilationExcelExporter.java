package ru.citlab24.protokol.tabs;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class VentilationExcelExporter {

    public static void export(List<VentilationRecord> records, java.awt.Component parent) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Вентиляция");

        // Стили
        CellStyle headerStyle = createHeaderStyle(workbook);
        CellStyle titleStyle = createTitleStyle(workbook);
        CellStyle dataStyle = createDataStyle(workbook);

        // Заголовок отчета
        createReportTitle(sheet, titleStyle);

        // Шапка таблицы
        createTableHeader(sheet, headerStyle);

        // Данные
        fillData(records, sheet, dataStyle);

        // Авторазмер колонок
        autoSizeColumns(sheet);

        // Сохранение файла
        saveWorkbook(workbook, parent);
    }

    private static CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 11);
        style.setFont(font);

        style.setWrapText(true);
        return style;
    }

    private static CellStyle createTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        style.setFont(font);
        return style;
    }

    private static CellStyle createDataStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.000"));
        return style;
    }

    private static void createReportTitle(Sheet sheet, CellStyle titleStyle) {
        Row titleRow = sheet.createRow(0);
        Cell titleCell = titleRow.createCell(0);
        titleCell.setCellValue("ПРОТОКОЛ ИЗМЕРЕНИЯ ПАРАМЕТРОВ ВЕНТИЛЯЦИОННЫХ СИСТЕМ");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

        Row dateRow = sheet.createRow(1);
        Cell dateCell = dateRow.createCell(0);
        dateCell.setCellValue("Дата: " + LocalDateTime.now().format(DateTimeFormatter.ofPattern("dd.MM.yyyy")));
    }

    private static void createTableHeader(Sheet sheet, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(3);
        String[] headers = {
                "№ п/п",
                "№точки измерения",
                "Рабочее место, место проведения измерений (Приток/вытяжка)",
                "Измеренные значения скорости воздушного потока (± расширенная неопределенность) м/с",
                "", // Пустая ячейка для объединения
                "", // Пустая ячейка для объединения
                "Площадь сечения проема, м²",
                "Производительность вент системы (канал или общая), м³/ч",
                "Объем помещения, м³",
                "Кратность воздухообмена по притоку или вытяжке, ч⁻¹",
                "Допустимый уровень производительности венсистем, м³/ч",
                "Допустимый уровень кратности воздухообмена, ч⁻¹"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }

        // Объединение ячеек для заголовка скорости
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 5));
        // Объединение пустых ячеек в D3-F3
    }

    private static void fillData(List<VentilationRecord> records, Sheet sheet, CellStyle dataStyle) {
        int rowNum = 4;
        int counter = 1;

        for (VentilationRecord record : records) {
            Row row = sheet.createRow(rowNum++);

            // № п/п
            Cell cell0 = row.createCell(0);
            cell0.setCellValue(counter++);
            cell0.setCellStyle(dataStyle);

            // № точки измерения (пусто)
            Cell cell1 = row.createCell(1);
            cell1.setCellStyle(dataStyle);

            // Место измерений
            Cell cell2 = row.createCell(2);
            cell2.setCellValue(String.format("Этаж %s, Помещение %s, %s",
                    record.floor(), record.space(), record.room()));
            cell2.setCellStyle(dataStyle);

            // Измеренные скорости (3 объединенные ячейки)
            Cell cell3 = row.createCell(3);
            cell3.setCellStyle(dataStyle);
            row.createCell(4).setCellStyle(dataStyle); // Пустые ячейки
            row.createCell(5).setCellStyle(dataStyle); // Пустые ячейки

            // Площадь сечения
            Cell cell6 = row.createCell(6);
            cell6.setCellValue(record.sectionArea());
            cell6.setCellStyle(dataStyle);

            // Производительность (пусто)
            Cell cell7 = row.createCell(7);
            cell7.setCellStyle(dataStyle);

            // Объем помещения
            Cell cell8 = row.createCell(8);
            cell8.setCellValue(record.volume() != null ? record.volume() : 0.0);
            cell8.setCellStyle(dataStyle);

            // Кратность (пусто)
            row.createCell(9).setCellStyle(dataStyle);
            // Допустимая производительность (пусто)
            row.createCell(10).setCellStyle(dataStyle);
            // Допустимая кратность (пусто)
            row.createCell(11).setCellStyle(dataStyle);
        }
    }

    private static void autoSizeColumns(Sheet sheet) {
        for (int i = 0; i < 12; i++) {
            sheet.autoSizeColumn(i);
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