package ru.citlab24.protokol.tabs;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.io.File;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

public class VentilationExcelExporter {

    public static void export(List<VentilationRecord> records, java.awt.Component parent) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Вентиляция");

        // Базовый шрифт Arial 10pt
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        // Стили
        CellStyle headerStyle = createHeaderStyle(workbook, baseFont);
        CellStyle titleStyle = createTitleStyle(workbook, baseFont);
        CellStyle dataStyle = createDataStyle(workbook, baseFont);
        CellStyle floorHeaderStyle = createFloorHeaderStyle(workbook, baseFont);

        // Создаем стили с разным форматированием чисел
        CellStyle twoDigitStyle = createNumberStyle(workbook, baseFont, "0.00"); // Два знака после запятой
        CellStyle integerStyle = createNumberStyle(workbook, baseFont, "0");     // Целое число
        CellStyle oneDigitStyle = createNumberStyle(workbook, baseFont, "0.0");  // Один знак после запятой

        // Создаем структуру документа
        createDocumentStructure(sheet, titleStyle);
        createTableHeaders(sheet, headerStyle);
        createColumnNumbers(sheet, headerStyle);

        // Устанавливаем ширину колонок
        setColumnsWidth(sheet);

        // Устанавливаем высоту строк
        setRowsHeight(sheet);

        fillData(records, sheet, dataStyle, twoDigitStyle, integerStyle, oneDigitStyle, floorHeaderStyle);

        // Сохранение файла
        saveWorkbook(workbook, parent);
    }

    private static void setColumnsWidth(Sheet sheet) {
        // Устанавливаем фиксированные ширины колонок
        sheet.setColumnWidth(0, 31 * 256 / 7);  // A: 31px
        sheet.setColumnWidth(1, 55 * 256 / 7);   // B: 55px
        sheet.setColumnWidth(2, 231 * 256 / 7);  // C: 231px
        sheet.setColumnWidth(3, 38 * 256 / 7);   // D: 38px
        sheet.setColumnWidth(4, 19 * 256 / 7);   // E: 19px
        sheet.setColumnWidth(5, 38 * 256 / 7);   // F: 38px
        sheet.setColumnWidth(6, 55 * 256 / 7);   // G: 55px
        sheet.setColumnWidth(7, 85 * 256 / 7);   // H: 85px
        sheet.setColumnWidth(8, 44 * 256 / 7);   // I: 44px
        sheet.setColumnWidth(9, 78 * 256 / 7);   // J: 78px
        sheet.setColumnWidth(10, 99 * 256 / 7);  // K: 99px
        sheet.setColumnWidth(11, 114 * 256 / 7); // L: 114px
    }

    private static void setRowsHeight(Sheet sheet) {
        // Устанавливаем высоты строк
        sheet.getRow(0).setHeightInPoints(12);     // 16px -> 12pt
        sheet.getRow(1).setHeightInPoints(6.75f);  // 9px -> 6.75pt
        sheet.getRow(2).setHeightInPoints(111);     // 148px -> 111pt

        // Строка 3 (номера колонок) - высота 15pt (20px)
        if (sheet.getRow(3) != null) {
            sheet.getRow(3).setHeightInPoints(15);
        }
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
        return style;
    }

    private static CellStyle createFloorHeaderStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        // Изменение: выравнивание по центру
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(baseFont);
        return style;
    }

    private static void createDocumentStructure(Sheet sheet, CellStyle titleStyle) {
        // Строка 0: Заголовок
        Row row0 = sheet.createRow(0);
        Cell titleCell = row0.createCell(0);
        titleCell.setCellValue("15.4. Скорость движения воздуха в рабочих проемах систем вентиляции, кратность воздухообмена");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

        // Строка 1: Пустая
        sheet.createRow(1);
    }

    private static void createTableHeaders(Sheet sheet, CellStyle headerStyle) {
        // Строка 2: Заголовки таблицы
        Row headerRow = sheet.createRow(2);
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
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 5));
    }

    private static void createColumnNumbers(Sheet sheet, CellStyle headerStyle) {
        // Строка 3: Номера колонок
        Row numberRow = sheet.createRow(3);

        for (int i = 0; i <= 11; i++) {
            Cell cell = numberRow.createCell(i);
            cell.setCellStyle(headerStyle);
        }

        // Заполняем значения
        numberRow.getCell(0).setCellValue("1");
        numberRow.getCell(1).setCellValue("2");
        numberRow.getCell(2).setCellValue("3");
        numberRow.getCell(3).setCellValue("4"); // Объединенная ячейка
        numberRow.getCell(6).setCellValue("5");
        numberRow.getCell(7).setCellValue("6");
        numberRow.getCell(8).setCellValue("7");
        numberRow.getCell(9).setCellValue("8");
        numberRow.getCell(10).setCellValue("9");
        numberRow.getCell(11).setCellValue("10");

        // Объединение для колонок D-F (индексы 3,4,5)
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 5));
    }

    private static void fillData(List<VentilationRecord> records, Sheet sheet,
                                 CellStyle dataStyle, CellStyle twoDigitStyle,
                                 CellStyle integerStyle, CellStyle oneDigitStyle,
                                 CellStyle floorHeaderStyle) {
        int rowNum = 4;
        int counter = 1;

        // Группируем записи по этажам
        Map<String, List<VentilationRecord>> recordsByFloor = records.stream()
                .collect(Collectors.groupingBy(
                        VentilationRecord::floor,
                        LinkedHashMap::new,
                        Collectors.toList()
                ));

        for (Map.Entry<String, List<VentilationRecord>> entry : recordsByFloor.entrySet()) {
            String floor = entry.getKey();
            List<VentilationRecord> floorRecords = entry.getValue();

            // Заголовок этажа (объединенная строка)
            Row floorRow = sheet.createRow(rowNum++);
            floorRow.setHeightInPoints(15); // 20px = 15pt

            Cell floorCell = floorRow.createCell(0);
            floorCell.setCellValue(floor);
            floorCell.setCellStyle(floorHeaderStyle);
            sheet.addMergedRegion(new CellRangeAddress(rowNum-1, rowNum-1, 0, 11));

            // Записи для каждого помещения на этаже
            for (VentilationRecord record : floorRecords) {
                Row dataRow = sheet.createRow(rowNum++);
                dataRow.setHeightInPoints(15); // 20px = 15pt

                // A: Порядковый номер
                Cell cell0 = dataRow.createCell(0);
                cell0.setCellValue(counter++);
                cell0.setCellStyle(dataStyle);

                // B: Всегда "-"
                Cell cell1 = dataRow.createCell(1);
                cell1.setCellValue("-");
                cell1.setCellStyle(dataStyle);

                // C: Наименование помещения и комнаты
                Cell cell2 = dataRow.createCell(2);
                cell2.setCellValue(record.space() + " " + record.room() + " (Вытяжка)");
                cell2.setCellStyle(dataStyle);

                // D: Расчетная скорость воздуха (два знака после запятой)
                double sectionArea = record.sectionArea();
                double targetFlow = 73; // Среднее между 65 и 81
                double airSpeed = targetFlow / (sectionArea * 3600);

                Cell cell3 = dataRow.createCell(3);
                cell3.setCellValue(airSpeed);
                cell3.setCellStyle(twoDigitStyle); // Формат: 0.00

                // E: Знак "±"
                Cell cell4 = dataRow.createCell(4);
                cell4.setCellValue("±");
                cell4.setCellStyle(dataStyle);

                // F: Формула неопределенности (два знака после запятой)
                Cell cell5 = dataRow.createCell(5);
                cell5.setCellFormula("(0.1+0.05*D" + rowNum + ")*2/(3^0.5)");
                cell5.setCellStyle(twoDigitStyle); // Формат: 0.00

                // G: Площадь сечения (два знака после запятой)
                Cell cell6 = dataRow.createCell(6);
                cell6.setCellValue(sectionArea);
                cell6.setCellStyle(twoDigitStyle); // Формат: 0.00

                // H: Производительность (целое число)
                Cell cell7 = dataRow.createCell(7);
                cell7.setCellFormula("ROUND(G" + rowNum + "*D" + rowNum + "*3600, 0)");
                cell7.setCellStyle(integerStyle); // Формат: 0

                // I: Объем помещения (один знак после запятой)
                Cell cell8 = dataRow.createCell(8);
                if (record.volume() != null && record.volume() > 0) {
                    cell8.setCellValue(record.volume());
                    cell8.setCellStyle(oneDigitStyle); // Формат: 0.0
                } else {
                    cell8.setCellValue("-");
                    cell8.setCellStyle(dataStyle);
                }

                // J: Кратность воздухообмена (один знак после запятой)
                Cell cell9 = dataRow.createCell(9);
                if (record.volume() != null && record.volume() > 0) {
                    cell9.setCellFormula("ROUND(H" + rowNum + "/I" + rowNum + ", 1)");
                    cell9.setCellStyle(oneDigitStyle); // Формат: 0.0
                } else {
                    cell9.setCellValue("-");
                    cell9.setCellStyle(dataStyle);
                }

                // K: Прочерк
                Cell cell10 = dataRow.createCell(10);
                cell10.setCellValue("-");
                cell10.setCellStyle(dataStyle);

                // L: Прочерк
                Cell cell11 = dataRow.createCell(11);
                cell11.setCellValue("-");
                cell11.setCellStyle(dataStyle);
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