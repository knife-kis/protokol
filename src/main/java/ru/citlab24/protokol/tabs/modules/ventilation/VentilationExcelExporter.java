package ru.citlab24.protokol.tabs.modules.ventilation;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.utils.RoomUtils;

import javax.swing.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ThreadLocalRandom;
import java.util.stream.Collectors;

public class VentilationExcelExporter {

    public static void export(List<VentilationRecord> records, java.awt.Component parent) {
        // Фильтруем записи: оставляем только те, у которых количество каналов > 0
        List<VentilationRecord> filteredRecords = records.stream()
                .filter(record -> record.channels() > 0)
                .collect(Collectors.toList());
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Вентиляция");

        // Базовый шрифт Arial 10pt
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        // Стили
        CellStyle headerStyle = createHeaderStyle(workbook, baseFont);
        CellStyle rotatedHeaderStyle = createRotatedHeaderStyle(workbook, baseFont);
        CellStyle titleStyle = createTitleStyle(workbook, baseFont);
        CellStyle dataStyle = createDataStyle(workbook, baseFont);
        CellStyle floorHeaderStyle = createFloorHeaderStyle(workbook, baseFont);
        CellStyle wrappedDataStyle = createWrappedDataStyle(workbook, baseFont);

        // Стили для группы D-F
        CellStyle plusMinusStyle = createPlusMinusStyle(workbook, baseFont);
        CellStyle leftInGroupStyle = createLeftInGroupStyle(workbook, baseFont);
        CellStyle rightInGroupStyle = createRightInGroupStyle(workbook, baseFont);

        // Стили с разным форматированием чисел
        CellStyle twoDigitStyle = createNumberStyle(workbook, baseFont, "0.00");
        CellStyle integerStyle = createNumberStyle(workbook, baseFont, "0");
        CellStyle oneDigitStyle = createNumberStyle(workbook, baseFont, "0.0");
        CellStyle threeDigitStyle = createNumberStyle(workbook, baseFont, "0.000");

        // Создаем структуру документа
        createDocumentStructure(sheet, titleStyle);
        createTableHeaders(sheet, headerStyle, rotatedHeaderStyle);
        createColumnNumbers(sheet, headerStyle, rotatedHeaderStyle);

        // Устанавливаем ширину колонок
        setColumnsWidth(sheet);

        // Устанавливаем высоту строк
        setRowsHeight(sheet);

        fillData(filteredRecords, sheet, dataStyle, wrappedDataStyle, threeDigitStyle, integerStyle, oneDigitStyle,
                floorHeaderStyle, plusMinusStyle, leftInGroupStyle, rightInGroupStyle);

        // Автоматическая высота строк для текста в столбце C
        adjustRowsWithText((XSSFSheet) sheet);

        // Сохранение файла
        saveWorkbook(workbook, parent);
    }

    private static void adjustRowsWithText(XSSFSheet sheet) {
        final float BASE_ROW_HEIGHT_POINTS = 15.0f; // 20 пикселей
        final float LARGE_TOTAL_HEIGHT_POINTS = 45.0f; // 60 пикселей
        final int COLUMN_C_WIDTH_PX = 231;
        final float CHAR_WIDTH_POINTS = 6.0f;

        for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cell = row.getCell(2);
            if (cell == null || cell.getCellType() != CellType.STRING) continue;

            String text = cell.getStringCellValue();
            if (text == null || text.isEmpty()) continue;

            // Определение объединенной области
            int mergedHeight = 1;
            boolean isFirstMerged = true;
            for (CellRangeAddress mergedRegion : sheet.getMergedRegions()) {
                if (mergedRegion.isInRange(rowIndex, 2)) {
                    mergedHeight = mergedRegion.getLastRow() - mergedRegion.getFirstRow() + 1;
                    isFirstMerged = (rowIndex == mergedRegion.getFirstRow());
                    break;
                }
            }
            if (!isFirstMerged) continue;

            // Расчет кол-ва строк текста
            int charsPerLine = (int) (COLUMN_C_WIDTH_PX / CHAR_WIDTH_POINTS);
            String[] lines = text.split("\r?\n");
            int lineCount = 0;
            for (String line : lines) {
                lineCount += Math.max(1, (int) Math.ceil((double) line.length() / charsPerLine));
            }

            // Определение общей высоты
            float totalHeightPoints;
            if (lineCount <= 1) {
                totalHeightPoints = BASE_ROW_HEIGHT_POINTS * mergedHeight;
            } else {
                totalHeightPoints = LARGE_TOTAL_HEIGHT_POINTS;
            }

            // Установка новой высоты
            float heightPerRow = totalHeightPoints / mergedHeight;
            for (int i = 0; i < mergedHeight; i++) {
                Row currentRow = sheet.getRow(rowIndex + i);
                if (currentRow == null) continue;
                currentRow.setHeightInPoints(heightPerRow);
            }
        }
    }

    private static void setRowsHeight(Sheet sheet) {
        sheet.getRow(0).setHeightInPoints(12);
        sheet.getRow(1).setHeightInPoints(6.75f);
        sheet.getRow(2).setHeightInPoints(111);
        if (sheet.getRow(3) != null) {
            sheet.getRow(3).setHeightInPoints(15); // 20 пикселей
        }
    }
    private static CellStyle createWrappedDataStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        style.setWrapText(true); // Включаем перенос текста
        setBlackBorderColor(style);
        return style;
    }
    private static CellStyle createRotatedHeaderStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        style.setWrapText(true);
        style.setRotation((short) 90);
        setBlackBorderColor(style);
        return style;
    }

    private static void setColumnsWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 31 * 256 / 7);
        sheet.setColumnWidth(1, 55 * 256 / 7);
        sheet.setColumnWidth(2, 231 * 256 / 7);
        sheet.setColumnWidth(3, 38 * 256 / 7);
        sheet.setColumnWidth(4, 19 * 256 / 7);
        sheet.setColumnWidth(5, 38 * 256 / 7);
        sheet.setColumnWidth(6, 55 * 256 / 7);
        sheet.setColumnWidth(7, 85 * 256 / 7);
        sheet.setColumnWidth(8, 44 * 256 / 7);
        sheet.setColumnWidth(9, 78 * 256 / 7);
        sheet.setColumnWidth(10, 99 * 256 / 7);
        sheet.setColumnWidth(11, 114 * 256 / 7);
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

    private static CellStyle createLeftInGroupStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.NONE);
        style.setFont(baseFont);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
        setBlackBorderColor(style);
        return style;
    }

    private static CellStyle createPlusMinusStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setFont(baseFont);
        setBlackBorderColor(style);
        return style;
    }

    private static CellStyle createRightInGroupStyle(Workbook workbook, Font baseFont) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.THIN);
        style.setFont(baseFont);
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));
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
        titleCell.setCellValue("15.4. Скорость движения воздуха в рабочих проемах систем вентиляции, кратность воздухообмена");
        titleCell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));
        sheet.createRow(1);
    }

    private static void createTableHeaders(Sheet sheet, CellStyle headerStyle, CellStyle rotatedHeaderStyle) {
        Row headerRow = sheet.createRow(2);
        String[] headers = {
                "№ п/п", "№точки измерения",
                "Рабочее место, место проведения измерений (Приток/вытяжка)",
                "Измеренные значения скорости воздушного потока (± расширенная неопределенность) м/с",
                "", "", "Площадь сечения проема, м²",
                "Производительность вент системы (канал или общая), м³/ч",
                "Объем помещения, м³", "Кратность воздухообмена по притоку или вытяжке, ч⁻¹",
                "Допустимый уровень производительности венсистем, м³/ч",
                "Допустимый уровень кратности воздухообмена, ч⁻¹"
        };

        for (int i = 0; i < headers.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(i >= 6 ? rotatedHeaderStyle : headerStyle);
        }
        sheet.addMergedRegion(new CellRangeAddress(2, 2, 3, 5));
    }

    private static void createColumnNumbers(Sheet sheet, CellStyle headerStyle, CellStyle rotatedHeaderStyle) {
        Row numberRow = sheet.createRow(3);
        for (int i = 0; i <= 11; i++) {
            Cell cell = numberRow.createCell(i);
            cell.setCellStyle(headerStyle);
        }
        numberRow.getCell(0).setCellValue("1");
        numberRow.getCell(1).setCellValue("2");
        numberRow.getCell(2).setCellValue("3");
        numberRow.getCell(3).setCellValue("4");
        numberRow.getCell(6).setCellValue("5");
        numberRow.getCell(7).setCellValue("6");
        numberRow.getCell(8).setCellValue("7");
        numberRow.getCell(9).setCellValue("8");
        numberRow.getCell(10).setCellValue("9");
        numberRow.getCell(11).setCellValue("10");
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 5));
    }

    private static void fillData(List<VentilationRecord> records, Sheet sheet,
                                 CellStyle dataStyle, CellStyle wrappedDataStyle,
                                 CellStyle threeDigitStyle, CellStyle integerStyle, CellStyle oneDigitStyle,
                                 CellStyle floorHeaderStyle, CellStyle plusMinusStyle,
                                 CellStyle leftInGroupStyle, CellStyle rightInGroupStyle) {
        int rowNum = 4;
        int counter = 1;

        Map<String, List<VentilationRecord>> recordsByFloor = records.stream()
                .collect(Collectors.groupingBy(
                        VentilationRecord::floor,
                        LinkedHashMap::new,
                        Collectors.toList()
                ));

        for (Map.Entry<String, List<VentilationRecord>> entry : recordsByFloor.entrySet()) {
            String floor = entry.getKey();
            List<VentilationRecord> floorRecords = entry.getValue();

            // Заголовок этажа
            Row floorRow = sheet.createRow(rowNum++);
            floorRow.setHeightInPoints(15); // 20 пикселей
            Cell floorCell = floorRow.createCell(0);
            floorCell.setCellValue(floor);
            floorCell.setCellStyle(floorHeaderStyle);
            CellRangeAddress mergedRegion = new CellRangeAddress(rowNum-1, rowNum-1, 0, 11);
            sheet.addMergedRegion(mergedRegion);
            RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegion, sheet);
            RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegion, sheet);

            for (VentilationRecord record : floorRecords) {
                int startRow = rowNum;
                int numChannels = record.channels();

                for (int i = 0; i < numChannels; i++) {
                    Row dataRow = sheet.createRow(rowNum++);
                    dataRow.setHeightInPoints(15);

                    // Колонка A
                    dataRow.createCell(0).setCellValue(counter++);
                    dataRow.getCell(0).setCellStyle(dataStyle);

                    // Колонка B
                    Cell cellB = dataRow.createCell(1);
                    cellB.setCellValue("-");
                    cellB.setCellStyle(dataStyle);

                    // Колонка C
                    if (i == 0) {
                        String displayName;
                        if (RoomUtils.isResidentialRoom(record.room())) {
                            String space = record.space();
                            String room = record.room();
                            String spaceNormalized = (space != null) ? space.trim().toLowerCase() : "";
                            String roomNormalized = (room != null) ? room.trim().toLowerCase() : "";

                            if (spaceNormalized.equals(roomNormalized)) {
                                displayName = space + " (Вытяжка)";
                            } else {
                                displayName = space + " " + room + " (Вытяжка)";
                            }
                        } else {
                            displayName = record.room() + " (Вытяжка)";
                        }
                        dataRow.createCell(2).setCellValue(displayName);
                        dataRow.getCell(2).setCellStyle(wrappedDataStyle);
                    } else {
                        dataRow.createCell(2).setCellValue("");
//                        dataRow.getCell(2).setCellStyle(dataStyle);
                    }
//                    dataRow.getCell(2).setCellStyle(dataStyle);

                    // Колонки D-F
                    double sectionArea = record.sectionArea();
                    double targetFlow = ThreadLocalRandom.current().nextInt(65, 82);
                    double airSpeed = targetFlow / (sectionArea * 3600);

                    Cell cellD = dataRow.createCell(3);
                    cellD.setCellValue(airSpeed);
                    cellD.setCellStyle(leftInGroupStyle);

                    Cell cellE = dataRow.createCell(4);
                    cellE.setCellValue("±");
                    cellE.setCellStyle(plusMinusStyle);

                    Cell cellF = dataRow.createCell(5);
                    cellF.setCellFormula("(0.1+0.05*D" + rowNum + ")*2/(3^0.5)");
                    cellF.setCellStyle(rightInGroupStyle);

                    // Колонка G
                    dataRow.createCell(6).setCellValue(sectionArea);
                    dataRow.getCell(6).setCellStyle(threeDigitStyle); // 3 знака после запятой

                    // Колонка H
                    Cell cellH = dataRow.createCell(7);
                    cellH.setCellFormula("ROUND(G" + rowNum + "*D" + rowNum + "*3600, 0)");
                    cellH.setCellStyle(integerStyle);

                    // Колонка I
                    if (i == 0) {
                        if (record.volume() != null && record.volume() > 0) {
                            dataRow.createCell(8).setCellValue(record.volume());
                            dataRow.getCell(8).setCellStyle(oneDigitStyle);
                        } else {
                            dataRow.createCell(8).setCellValue("-");
                            dataRow.getCell(8).setCellStyle(dataStyle);
                        }
                    } else {
                        dataRow.createCell(8).setCellValue("");
                        dataRow.getCell(8).setCellStyle(dataStyle);
                    }

                    // Колонка J
                    dataRow.createCell(9);

                    // Колонки K-L
                    if (i == 0) {
                        Double airExchangeRate = RoomUtils.getAirExchangeRate(record.room());

                        // Колонка K
                        Cell cellK = dataRow.createCell(10);
                        if (airExchangeRate != null) {
                            cellK.setCellValue(Math.round(airExchangeRate)); // Целое число
                            cellK.setCellStyle(integerStyle);
                        } else {
                            cellK.setCellValue("-");
                            cellK.setCellStyle(dataStyle);
                        }

                        // Колонка L - всегда "-"
                        Cell cellL = dataRow.createCell(11);
                        cellL.setCellValue("-");
                        cellL.setCellStyle(dataStyle);
                    }
                }

                int lastRow = rowNum - 1;

                // Заполнение колонки J
                if (record.volume() != null && record.volume() > 0) {
                    Row firstRow = sheet.getRow(startRow);
                    Cell cellJ = firstRow.getCell(9);

                    if (numChannels == 1) {
                        cellJ.setCellFormula("ROUND(H" + (startRow+1) + "/I" + (startRow+1) + ", 1)");
                    } else {
                        StringBuilder sumFormula = new StringBuilder("SUM(H" + (startRow+1));
                        for (int i = startRow+2; i <= lastRow+1; i++) {
                            sumFormula.append(",H").append(i);
                        }
                        sumFormula.append(")");
                        cellJ.setCellFormula("ROUND(" + sumFormula + "/I" + (startRow+1) + ", 1)");
                    }
                    cellJ.setCellStyle(oneDigitStyle);
                } else {
                    Row firstRow = sheet.getRow(startRow);
                    Cell cellJ = firstRow.getCell(9);
                    cellJ.setCellValue("-");
                    cellJ.setCellStyle(dataStyle);
                }

                // Объединение ячеек с восстановлением границ
                int[] columnsToMerge = {2, 8, 9, 10, 11};
                if (lastRow > startRow) {
                    for (int col : columnsToMerge) {
                        CellRangeAddress mergedRegionRecord = new CellRangeAddress(startRow, lastRow, col, col);
                        sheet.addMergedRegion(mergedRegionRecord);

                        // Восстановление границ
                        RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegionRecord, sheet);
                        RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegionRecord, sheet);
                        RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegionRecord, sheet);
                        RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegionRecord, sheet);
                    }
                }
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