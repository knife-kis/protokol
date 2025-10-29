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

    // ВНИМАНИЕ: это новый метод, просто вставь в класс VentilationExcelExporter.
    public static void appendToWorkbook(java.util.List<VentilationRecord> records, org.apache.poi.ss.usermodel.Workbook wb) {
        if (wb == null || records == null || records.isEmpty()) return;

        // 1) Имя листа без коллизий
        String baseName = "Вентиляция";
        String sheetName = baseName;
        int i = 2;
        while (wb.getSheet(sheetName) != null) sheetName = baseName + " (" + (i++) + ")";

        org.apache.poi.ss.usermodel.Sheet sheet = wb.createSheet(sheetName);

        // 2) Базовый шрифт
        org.apache.poi.ss.usermodel.Font baseFont = wb.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        // 3) Стили (используем уже имеющиеся фабрики)
        org.apache.poi.ss.usermodel.CellStyle headerStyle        = createHeaderStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle rotatedHeaderStyle = createRotatedHeaderStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle titleStyle         = createTitleStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle dataStyle          = createDataStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle floorHeaderStyle   = createFloorHeaderStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle wrappedDataStyle   = createWrappedDataStyle(wb, baseFont);

        org.apache.poi.ss.usermodel.CellStyle plusMinusStyle     = createPlusMinusStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle leftInGroupStyle   = createLeftInGroupStyle(wb, baseFont);
        org.apache.poi.ss.usermodel.CellStyle rightInGroupStyle  = createRightInGroupStyle(wb, baseFont);

        org.apache.poi.ss.usermodel.CellStyle twoDigitStyle   = createNumberStyle(wb, baseFont, "0.00");
        org.apache.poi.ss.usermodel.CellStyle integerStyle    = createNumberStyle(wb, baseFont, "0");
        org.apache.poi.ss.usermodel.CellStyle oneDigitStyle   = createNumberStyle(wb, baseFont, "0.0");
        org.apache.poi.ss.usermodel.CellStyle threeDigitStyle = createNumberStyle(wb, baseFont, "0.000");

        // 4) Шапка/структура/габариты
        createDocumentStructure(sheet, titleStyle);
        createTableHeaders(sheet, headerStyle, rotatedHeaderStyle);
        createColumnNumbers(sheet, headerStyle, rotatedHeaderStyle);
        setColumnsWidth(sheet);
        setRowsHeight(sheet);

        // 5) Данные (как в export, только без сохранения)
        java.util.List<VentilationRecord> filtered = records.stream()
                .filter(r -> r != null && r.channels() > 0)
                .collect(java.util.stream.Collectors.toList());

        fillData(filtered, sheet, dataStyle, wrappedDataStyle, threeDigitStyle, integerStyle, oneDigitStyle,
                floorHeaderStyle, plusMinusStyle, leftInGroupStyle, rightInGroupStyle);

        // 6) Автовысота строк для текста в C
        adjustRowsWithText((org.apache.poi.xssf.usermodel.XSSFSheet) sheet);
    }

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
        wrappedDataStyle.setWrapText(true);

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
        final float BASE_ROW_HEIGHT_POINTS = 15.0f; // базовая высота одной строки
        final int COLUMN_INDEX = 2;                 // столбец C

        // ширина колонки C в "символах" (Excel хранит в 1/256 символа)
        final int colWidth256 = sheet.getColumnWidth(COLUMN_INDEX);
        final int charsPerLine = Math.max(1, colWidth256 / 256);

        for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;

            Cell cell = row.getCell(COLUMN_INDEX);
            if (cell == null || cell.getCellType() != CellType.STRING) continue;

            String text = cell.getStringCellValue();
            if (text == null || text.isEmpty()) continue;

            // Определяем объединение по столбцу C
            int mergedHeight = 1;
            boolean isFirstMerged = true;
            for (CellRangeAddress mr : sheet.getMergedRegions()) {
                if (mr.isInRange(rowIndex, COLUMN_INDEX)) {
                    mergedHeight = mr.getLastRow() - mr.getFirstRow() + 1;
                    isFirstMerged = (rowIndex == mr.getFirstRow());
                    break;
                }
            }
            // Размер меняем только в первой строке объединения
            if (!isFirstMerged) continue;

            // Оценка требуемого количества «визуальных строк» с учётом wrapText
            int neededLines = 0;
            for (String para : text.split("\\R", -1)) {
                if (para.isEmpty()) { neededLines++; continue; }
                neededLines += (para.length() + charsPerLine - 1) / charsPerLine; // ceil(len / charsPerLine)
            }

            // «Добавляем высоту одного ряда, пока не влезет»:
            // нужно не меньше, чем mergedHeight, и не меньше, чем neededLines.
            int totalVisualLines = Math.max(mergedHeight, neededLines);

            // Распределяем общую высоту на все строки объединения
            float heightPerRow = (BASE_ROW_HEIGHT_POINTS * totalVisualLines) / mergedHeight;
            for (int i = 0; i < mergedHeight; i++) {
                Row r = sheet.getRow(rowIndex + i);
                if (r != null) r.setHeightInPoints(heightPerRow);
            }
        }
    }

    // Оценка числа «визуальных строк» для текста в колонке при wrapText=true.
// Берём ширину колонки в "символах" (Excel хранит в 1/256 символа), делим текст по переносам
// и считаем, сколько «строк» потребуется с учётом ширины.
    private static int estimateLinesForCell(Sheet sheet, int columnIndex, String text) {
        if (text == null || text.isBlank()) return 1;

        // ширина колонки в символах (очень близкая к тому, как считает Excel autosize)
        int colWidth256 = sheet.getColumnWidth(columnIndex); // в 1/256 символа
        int colChars = Math.max(1, colWidth256 / 256);

        int lines = 0;
        for (String para : text.split("\\R", -1)) {
            if (para.isEmpty()) { lines++; continue; }
            // грубая оценка: длинные слова тоже считаем по символам
            lines += (para.length() + colChars - 1) / colChars;
        }
        return Math.max(1, lines);
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
        final float BASE_ROW_HEIGHT = Math.max(15f, sheet.getDefaultRowHeightInPoints());

        // Нужно ли показывать "блок-секцию" в шапке (только если секций > 1)
        int distinctSections = (int) records.stream()
                .map(VentilationRecord::sectionIndex)
                .distinct()
                .count();
        boolean showSectionInHeader = distinctSections > 1;

        // Группировка: секция (по возрастанию индекса) → этаж (в порядке появления)
        java.util.Map<Integer, java.util.Map<String, java.util.List<VentilationRecord>>> bySectionThenFloor =
                records.stream().collect(
                        java.util.stream.Collectors.groupingBy(
                                VentilationRecord::sectionIndex,
                                java.util.TreeMap::new, // секции 0,1,2,...
                                java.util.stream.Collectors.groupingBy(
                                        VentilationRecord::floor,
                                        java.util.LinkedHashMap::new,
                                        java.util.stream.Collectors.toList()
                                )
                        )
                );

        for (java.util.Map.Entry<Integer, java.util.Map<String, java.util.List<VentilationRecord>>> secEntry
                : bySectionThenFloor.entrySet()) {
            int sectionIndex = secEntry.getKey();
            java.util.Map<String, java.util.List<VentilationRecord>> byFloor = secEntry.getValue();

            for (java.util.Map.Entry<String, java.util.List<VentilationRecord>> entry : byFloor.entrySet()) {
                String floor = entry.getKey();
                java.util.List<VentilationRecord> floorRecords = entry.getValue();

                // Заголовок: "блок-секция N, X этаж" либо просто "X этаж" если секция одна
                String headerText = showSectionInHeader
                        ? ("блок-секция " + (sectionIndex + 1) + ", " + floor)
                        : floor;

                Row floorRow = sheet.createRow(rowNum++);
                floorRow.setHeightInPoints(15);
                Cell floorCell = floorRow.createCell(0);
                floorCell.setCellValue(headerText);
                floorCell.setCellStyle(floorHeaderStyle);
                CellRangeAddress mergedRegion = new CellRangeAddress(rowNum - 1, rowNum - 1, 0, 11);
                sheet.addMergedRegion(mergedRegion);
                RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegion, sheet);
                RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegion, sheet);
                RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegion, sheet);
                RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegion, sheet);

                // ===== ниже — твоя текущая логика заполнения строк (без изменений) =====
                for (VentilationRecord record : floorRecords) {
                    int startRow = rowNum;
                    int numChannels = record.channels();

                    for (int i = 0; i < numChannels; i++) {
                        Row dataRow = sheet.createRow(rowNum++);

                        // A: № п/п
                        dataRow.createCell(0).setCellValue(counter++);
                        dataRow.getCell(0).setCellStyle(dataStyle);

                        // B: место отбора пробы (пока "-")
                        Cell cellB = dataRow.createCell(1);
                        cellB.setCellValue("-");
                        cellB.setCellStyle(dataStyle);

                        // C: место измерения (группа по каналам)
                        if (i == 0) {
                            String displayName;
                            if (RoomUtils.isResidentialRoom(record.room())) {
                                String space = record.space();
                                String roomName = record.room();
                                displayName = (space == null || space.isBlank() ? roomName : space + ", " + roomName) + " (Вытяжка)";
                            } else {
                                displayName = record.room() + " (Вытяжка)";
                            }
                            Cell cellC = dataRow.createCell(2);
                            cellC.setCellValue(displayName);
                            cellC.setCellStyle(wrappedDataStyle);

// === авто-подбор высоты для столбца C ===
// Сколько визуальных строк требуется при текущей ширине колонки C:
                            int neededLines = estimateLinesForCell(sheet, 2, displayName);

// Сколько строк доступно «по умолчанию»: это число каналов (т.к. C объединяем по группе).
// Если канал один — доступна 1 строка; если больше — доступна сумма высот всех строк группы.
                            int availableLines = Math.max(1, numChannels);

// Если требуется больше, чем доступно, добавляем недостающее к высоте первой строки группы.
// Таким образом суммарная высота объединённой области C будет достаточной.
                            int totalLines = Math.max(neededLines, availableLines);
                            float totalHeightPts = BASE_ROW_HEIGHT * totalLines;
                            float firstRowHeightPts = totalHeightPts - BASE_ROW_HEIGHT * (availableLines - 1);

// На случай погрешностей округляем вверх на пол-пункта
                            dataRow.setHeightInPoints((float)Math.ceil(firstRowHeightPts + 0.5));

                        } else {
                            Cell cellC = dataRow.createCell(2);
                            cellC.setCellValue("");
                            cellC.setCellStyle(wrappedDataStyle);
                        }

                        // D–F: скорость, ±, неопределенность (как у тебя было)
                        double sectionArea = record.sectionArea();
                        double targetFlow = java.util.concurrent.ThreadLocalRandom.current().nextInt(65, 82);
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

                        // G: сечение
                        dataRow.createCell(6).setCellValue(sectionArea);
                        dataRow.getCell(6).setCellStyle(threeDigitStyle);

                        // H: производительность
                        Cell cellH = dataRow.createCell(7);
                        cellH.setCellFormula("ROUND(G" + rowNum + "*D" + rowNum + "*3600, 0)");
                        cellH.setCellStyle(integerStyle);

                        // I: объем помещения — только в первой строке группы
                        // Колонка I (объем помещения) — только в первой строке группы
                        if (i == 0) {
                            Cell cellI = dataRow.createCell(8);
                            if (record.volume() != null && record.volume() > 0) {
                                cellI.setCellValue(record.volume());
                                cellI.setCellStyle(oneDigitStyle);
                            } else {
                                cellI.setCellValue("-");           // ← ставим дефолтное тире
                                cellI.setCellStyle(dataStyle);     // стиль для текста/центра
                            }
                        } else {
                            // нижние строки объединения — ставим дефолтный тире
                            Cell cellI = dataRow.createCell(8);
                            cellI.setCellValue("-");
                            cellI.setCellStyle(dataStyle);
                        }


                        // J: кратность (формула заполняется после группы)
                        Cell cellJ = dataRow.createCell(9);
                        cellJ.setCellValue("-");
                        cellJ.setCellStyle(dataStyle);

                        // K–L: нормы/допустимые
                        if (i == 0) {
                            Double airExchangeRate = RoomUtils.getAirExchangeRate(record.room());

                            // K
                            Cell cellK = dataRow.createCell(10);
                            if (airExchangeRate != null) {
                                cellK.setCellValue(Math.round(airExchangeRate));
                                cellK.setCellStyle(integerStyle);
                            } else {
                                cellK.setCellValue("-");
                                cellK.setCellStyle(dataStyle);
                            }

                            // L — базово "-"
                            Cell cellL = dataRow.createCell(11);
                            cellL.setCellValue("-");
                            cellL.setCellStyle(dataStyle);
                        } else {
                            // нижние строки объединения — тоже ставим дефолтные тире
                            Cell cellK = dataRow.createCell(10);
                            cellK.setCellValue("-");
                            cellK.setCellStyle(dataStyle);

                            Cell cellL = dataRow.createCell(11);
                            cellL.setCellValue("-");
                            cellL.setCellStyle(dataStyle);
                        }
                    }

                    int lastRow = rowNum - 1;

                    // Заполнение колонки J (кратность) одной формулой по группе
                    if (record.volume() != null && record.volume() > 0) {
                        Row firstRow = sheet.getRow(startRow);
                        Cell cellJ = firstRow.getCell(9);

                        if (numChannels == 1) {
                            cellJ.setCellFormula("ROUND(H" + (startRow + 1) + "/I" + (startRow + 1) + ", 1)");
                        } else {
                            StringBuilder sumFormula = new StringBuilder("SUM(H" + (startRow + 1));
                            for (int i = startRow + 2; i <= lastRow + 1; i++) {
                                sumFormula.append(",H").append(i);
                            }
                            sumFormula.append(")");
                            cellJ.setCellFormula("ROUND(" + sumFormula + "/I" + (startRow + 1) + ", 1)");
                        }
                        cellJ.setCellStyle(oneDigitStyle);
                    } else {
                        Row firstRow = sheet.getRow(startRow);
                        Cell cellJ = firstRow.getCell(9);
                        cellJ.setCellValue("-");
                        cellJ.setCellStyle(dataStyle);
                    }

                    // Объединение C, I, J, K, L по группе (если каналов > 1) с восстановлением границ
                    int[] columnsToMerge = {2, 8, 9, 10, 11};
                    if (lastRow > startRow) {
                        for (int col : columnsToMerge) {
                            CellRangeAddress mergedRegionRecord = new CellRangeAddress(startRow, lastRow, col, col);
                            sheet.addMergedRegion(mergedRegionRecord);
                            RegionUtil.setBorderTop(BorderStyle.THIN, mergedRegionRecord, sheet);
                            RegionUtil.setBorderBottom(BorderStyle.THIN, mergedRegionRecord, sheet);
                            RegionUtil.setBorderLeft(BorderStyle.THIN, mergedRegionRecord, sheet);
                            RegionUtil.setBorderRight(BorderStyle.THIN, mergedRegionRecord, sheet);
                        }
                    }
                }
                // ===== конец блока заполнения строк =====
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