package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateExcelExporter;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationExcelExporter;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationRecord;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationTab;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.util.List;


public final class AllExcelExporter {
    private AllExcelExporter() {}

    // Стало
    public static void exportAll(MainFrame frame, Building building, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = new XSSFWorkbook()) {
            // 0) Титульная страница — первым листом
            appendTitleSheet(frame, building, wb);

            // 1) Микроклимат
            MicroclimateExcelExporter.appendToWorkbook(building, -1, wb);

            // 2) Вентиляция (берём текущие записи из вкладки)
            List<VentilationRecord> ventRecords = null;
            if (frame != null && frame.getVentilationTab() != null) {
                ventRecords = frame.getVentilationTab().getRecordsForExport();
            }
            if (ventRecords != null && !ventRecords.isEmpty()) {
                VentilationExcelExporter.appendToWorkbook(ventRecords, wb);
            }

            // 3) «Осв улица» — перед КЕО
            appendStreetLightingSheet(wb, building);

            // 4) Естественное освещение (КЕО)
            tryInvokeAppend("ru.citlab24.protokol.tabs.modules.lighting.LightingExcelExporter",
                    new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                    new Object[]{building, -1, wb});

            // 5) Искусственное освещение
            java.util.Map<Integer, Boolean> litMap = null;
            try {
                if (frame != null && frame.getArtificialLightingTab() != null) {
                    litMap = frame.getArtificialLightingTab().snapshotSelectionMap();
                }
            } catch (Throwable ignore) {}
            try {
                Class<?> clazz = Class.forName("ru.citlab24.protokol.tabs.modules.lighting.ArtificialLightingExcelExporter");
                java.lang.reflect.Method m = clazz.getMethod(
                        "appendToWorkbook",
                        ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class, java.util.Map.class
                );
                m.invoke(null, building, -1, wb, litMap);
            } catch (NoSuchMethodException e) {
                tryInvokeAppend("ru.citlab24.protokol.tabs.modules.lighting.ArtificialLightingExcelExporter",
                        new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                        new Object[]{building, -1, wb});
            } catch (Throwable t) {
                t.printStackTrace();
            }

            // 6) Радиация
            tryInvokeAppend("ru.citlab24.protokol.tabs.modules.med.RadiationExcelExporter",
                    new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                    new Object[]{building, -1, wb});

            // Приводим ВСЕ листы к печати "в 1 страницу по ширине"
            applyFitToPageWidthForAllSheets(wb);

            // 7) Диалог сохранения
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить общий Excel");
            chooser.setSelectedFile(new File("Отчет_все_модули.xlsx"));
            if (chooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = chooser.getSelectedFile();
                if (!file.getName().toLowerCase().endsWith(".xlsx")) {
                    file = new File(file.getAbsolutePath() + ".xlsx");
                }
                try (FileOutputStream out = new FileOutputStream(file)) {
                    wb.write(out);
                }
                JOptionPane.showMessageDialog(parent,
                        "Файл сохранён:\n" + file.getAbsolutePath(),
                        "Экспорт завершён", JOptionPane.INFORMATION_MESSAGE);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка экспорта: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }
    /** Первый лист: титульная страница. */
    private static void appendTitleSheet(MainFrame frame, Building building, Workbook wb) {
        Sheet sheet = wb.createSheet("Титульная страница");

        // Ориентация: альбомная, А4
        PrintSetup ps = sheet.getPrintSetup();
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);
        sheet.setFitToPage(true);

        // Базовый стиль: Arial 10, перенос по словам
        Font baseFont = wb.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        CellStyle baseStyle = wb.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setVerticalAlignment(VerticalAlignment.TOP);
        baseStyle.setAlignment(HorizontalAlignment.LEFT);

        CellStyle rightStyle = wb.createCellStyle();
        rightStyle.cloneStyleFrom(baseStyle);
        rightStyle.setAlignment(HorizontalAlignment.RIGHT);

        CellStyle centerStyle = wb.createCellStyle();
        centerStyle.cloneStyleFrom(baseStyle);
        centerStyle.setAlignment(HorizontalAlignment.CENTER);

        // Высоты строк оставляем как раньше — тебе они подошли
        for (int r = 0; r <= 10; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            row.setHeightInPoints(16f);
        }

        // Ширины столбцов — берём из эталонного файла (величины в "символах" Excel).
        // Подобрано так, чтобы столбец B давал около 37 px по ширине.
        // A..Z => индексы 0..25
        double[] colCharWidths = new double[] {
                5.14,      // A
                5.14,      // B (≈ 37 px)
                5.14,     // C
                5.14,      // D
                5.14,      // E
                5.14,      // F
                5.14,      // G
                5.14,      // H
                5.14,      // I
                5.14,      // J
                5.14,      // K
                5.14,      // L
                5.14,      // M
                5.14,      // N
                5.14,      // O
                5.14,      // P
                5.14,      // Q
                5.14,      // R
                5.14,      // S
                5.14,      // T
                5.14,      // U
                5.14,      // V
                5.14,  // W
                5.14,     // X
                5.14,      // Y
                7.14       // Z
        };

        for (int i = 0; i < colCharWidths.length; i++) {
            double w = colCharWidths[i];
            if (w > 0) {
                sheet.setColumnWidth(i, (int) Math.round(w * 256));
            }
        }

        // 1) A1-I3
        String txtA1 = "Общество с ограниченной ответственностью «Центр исследовательских технологий»\n" +
                "(ООО «ЦИТ»)";
        setMergedText(sheet, baseStyle, 0, 2, 0, 8, txtA1);

        // 2) S1-Z3
        String txtS1 = "УТВЕРЖДАЮ\nЗаместитель заведующего лабораторией";
        setMergedText(sheet, rightStyle, 0, 2, 18, 25, txtS1);

        // 3) A4-I5
        String txtA4 = "660064, Красноярский край, г.о. Красноярск,\n" +
                "г. Красноярск, ул. Парусная, д. 8, кв. 99";
        setMergedText(sheet, baseStyle, 3, 4, 0, 8, txtA4);

        // 4) A6-I8
        String txtA6 = "Испытательная лаборатория\n" +
                "Общества с ограниченной ответственностью\n" +
                "«Центр исследовательских технологий»\n" +
                "Уникальный номер записи в Реестре аккредитованных лиц RA.RU.21ОХ37 от 07.04.2023";
        setMergedText(sheet, baseStyle, 5, 7, 0, 8, txtA6);

        // 5) A9-I11
        String txtA9 = "660041, Красноярский край, г. Красноярск,\n" +
                "пр. Свободный дом 75, пом. 9 кабинет 12\n" +
                "+7 (391) 244-03-10, cit-krsk@yandex.ru";
        setMergedText(sheet, baseStyle, 8, 10, 0, 8, txtA9);

        // S5-Z5: _______________/М.Е. Гаврилова/
        String txtS5 = "_______________/М.Е. Гаврилова/";
        setMergedText(sheet, rightStyle, 4, 4, 18, 25, txtS5);

        // S7-Z7: дата из программы + " г."
        String protocolDate = resolveProtocolDateText(frame);
        String txtS7;
        if (protocolDate != null && !protocolDate.isBlank()) {
            txtS7 = protocolDate.trim() + " г.";
        } else {
            // запасной вариант, если датy из программы не нашли
            txtS7 = "____.__.____ г.";
        }
        setMergedText(sheet, rightStyle, 6, 6, 18, 25, txtS7);

        // S9-Z9: МП
        setMergedText(sheet, centerStyle, 8, 8, 18, 25, "МП");
    }

    private static void tryInvokeAppend(String fqcn, Class<?>[] sig, Object[] args) {
        try {
            Class<?> clazz = Class.forName(fqcn);
            Method m = clazz.getMethod("appendToWorkbook", sig);
            m.invoke(null, args); // static
        } catch (ClassNotFoundException ignored) {
            // класса нет — тихо пропускаем
        } catch (NoSuchMethodException ignored) {
            // метода appendToWorkbook пока нет — пропускаем
        } catch (Throwable t) {
            // любая другая ошибка — пишем в консоль, но не валим общий экспорт
            t.printStackTrace();
        }
    }
    // === НОВОЕ: «Осв улица» — собираем строки и добавляем лист в текущую книгу ===
    private static void appendStreetLightingSheet(Workbook wb, Building building) {
        try {
            // 1) Подтянуть сохранённые 4 значения по ключам (если есть)
            java.util.Map<String, Double[]> byKey = java.util.Collections.emptyMap();
            try {
                byKey = ru.citlab24.protokol.db.DatabaseManager
                        .loadStreetLightingValuesByKey(building.getId());
            } catch (java.sql.SQLException ex) {
                System.err.println("[AllExcelExporter] WARN: не удалось прочитать 'Осв улица' из БД: " + ex.getMessage());
            }

            // 2) Собрать все комнаты на этажах STREET → помещения только OUTDOOR
            java.util.List<ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.RowData> rows =
                    new java.util.ArrayList<>();

            java.util.List<ru.citlab24.protokol.tabs.models.Floor> floors =
                    new java.util.ArrayList<>(building.getFloors());
            floors.sort(java.util.Comparator.comparingInt(
                    ru.citlab24.protokol.tabs.models.Floor::getPosition));

            for (ru.citlab24.protokol.tabs.models.Floor f : floors) {
                if (f == null || f.getType() != ru.citlab24.protokol.tabs.models.Floor.FloorType.STREET) continue;

                java.util.List<ru.citlab24.protokol.tabs.models.Space> spaces =
                        new java.util.ArrayList<>(f.getSpaces());
                spaces.sort(java.util.Comparator.comparingInt(
                        ru.citlab24.protokol.tabs.models.Space::getPosition));

                for (ru.citlab24.protokol.tabs.models.Space s : spaces) {
                    if (s == null || s.getType() != ru.citlab24.protokol.tabs.models.Space.SpaceType.OUTDOOR) continue;

                    java.util.List<ru.citlab24.protokol.tabs.models.Room> rooms =
                            new java.util.ArrayList<>(s.getRooms());
                    rooms.sort(java.util.Comparator.comparingInt(
                            ru.citlab24.protokol.tabs.models.Room::getPosition));

                    for (ru.citlab24.protokol.tabs.models.Room r : rooms) {
                        if (r == null) continue;

                        String floorPart = (f.getNumber() == null) ? "" : f.getNumber().trim();
                        String spacePart = (s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
                        String roomPart  = (r.getName() == null) ? "" : r.getName().trim();

                        // Ключ как во вкладке StreetLightingTab:
                        // sectionIndex|этаж|помещение|комната
                        String key = f.getSectionIndex() + "|" + floorPart + "|" + spacePart + "|" + roomPart;

                        Double leftMax = null, centerMin = null, rightMax = null, bottomMin = null;
                        Double[] vals = (byKey != null) ? byKey.get(key) : null;
                        if (vals != null) {
                            if (vals.length > 0) leftMax   = vals[0];
                            if (vals.length > 1) centerMin = vals[1];
                            if (vals.length > 2) rightMax  = vals[2];
                            if (vals.length > 3) bottomMin = vals[3];
                        }

                        rows.add(new ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.RowData(
                                roomPart, leftMax, centerMin, rightMax, bottomMin
                        ));
                    }
                }
            }

            // 3) Вставить лист «Осв улица» в текущую книгу
            ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.appendToWorkbook(rows, wb);

            // 4) Попробовать зафиксировать порядок: «Осв улица» перед «Естественное освещение»
            try {
                int streetIdx = wb.getSheetIndex("Осв улица");
                int keoIdx    = wb.getSheetIndex("Естественное освещение");
                if (streetIdx >= 0 && keoIdx >= 0 && streetIdx > keoIdx) {
                    wb.setSheetOrder("Осв улица", keoIdx);
                }
            } catch (Throwable ignore) {
                // на случай, если лист КЕО назван иначе — порядок уже обеспечен самим вызовом до КЕО
            }
        } catch (Throwable t) {
            System.err.println("[AllExcelExporter] Ошибка добавления 'Осв улица': " + t.getMessage());
        }
    }
    /** Применяет настройку печати "уместить по ширине в 1 страницу" ко всем листам книги. */
    private static void applyFitToPageWidthForAllSheets(Workbook wb) {
        if (wb == null) return;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(i);
            if (sh != null) tuneSheetToOnePageWidth(sh);
        }
    }
    /**
     * Ширины столбцов и высоты строк для титульного листа
     * скопированы из файла 29_10.07.25-1-Ф_25.xlsm, лист "Протокол".
     */
    private static void applyTitleSheetGrid(Sheet sheet) {
        // === ШИРИНЫ СТОЛБЦОВ A..Z ===
        // значения в единицах POI: 1/256 символа
        // A = 5.00; B..Y = 13.00; Z = 8.14
        int[] colWidths = new int[] {
                1280, // A (5.00)
                3328, // B (13.00)
                3328, // C
                3328, // D
                3328, // E
                3328, // F
                3328, // G
                3328, // H
                3328, // I
                3328, // J
                3328, // K
                3328, // L
                3328, // M
                3328, // N
                3328, // O
                3328, // P
                3328, // Q
                3328, // R
                3328, // S
                3328, // T
                3328, // U
                3328, // V
                3328, // W
                3328, // X
                3328, // Y
                2084  // Z (8.14)
        };
        for (int c = 0; c < colWidths.length; c++) {
            sheet.setColumnWidth(c, colWidths[c]);
        }

        // === ВЫСОТЫ СТРОК 1–11 (в пунктах) ===
        // 1: 15.0
        // 2: 15.75
        // 3: 7.5
        // 4: (по умолчанию 15.0)
        // 5: 12.0
        // 6: (по умолчанию 15.0)
        // 7: (по умолчанию 15.0)
        // 8: 35.25
        // 9: 15.0
        // 10: (по умолчанию 15.0)
        // 11: 14.25
        float[] rowHeights = new float[] {
                15.0f,   // row 1
                15.75f,  // row 2
                7.5f,    // row 3
                15.0f,   // row 4
                12.0f,   // row 5
                15.0f,   // row 6
                15.0f,   // row 7
                35.25f,  // row 8
                15.0f,   // row 9
                15.0f,   // row 10
                14.25f   // row 11
        };

        for (int r = 0; r < rowHeights.length; r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                row = sheet.createRow(r);
            }
            row.setHeightInPoints(rowHeights[r]);
        }
    }

    /**
     * Пытаемся вытащить дату протокола из программы:
     * 1) метод на MainFrame: getProtocolDateForExcel() / getProtocolDateString() / getTitlePageProtocolDate()
     * 2) либо на вкладке TitlePageTab: метод getProtocolDateForExcel()
     *
     * Если ничего не нашли — возвращаем null.
     */
    private static String resolveProtocolDateText(MainFrame frame) {
        if (frame == null) return null;

        // 1) Прямые геттеры на MainFrame (если ты их сделаешь)
        for (String methodName : new String[]{
                "getProtocolDateForExcel",
                "getProtocolDateString",
                "getTitlePageProtocolDate"
        }) {
            try {
                Method m = frame.getClass().getMethod(methodName);
                Object val = m.invoke(frame);
                if (val != null) {
                    String s = val.toString().trim();
                    if (!s.isEmpty()) return s;
                }
            } catch (NoSuchMethodException ignored) {
                // метода нет — просто пробуем следующий
            } catch (Throwable ignored) {
                // любые другие ошибки — тихо игнорируем
            }
        }

        // 2) Ищем вкладку с классом, в имени которого есть "TitlePageTab"
        try {
            Method mTabs = frame.getClass().getMethod("getTabbedPane");
            Object tabsObj = mTabs.invoke(frame);
            if (tabsObj instanceof JTabbedPane tabs) {
                int count = tabs.getTabCount();
                for (int i = 0; i < count; i++) {
                    Component c = tabs.getComponentAt(i);
                    if (c == null) continue;
                    Class<?> clazz = c.getClass();
                    if (clazz.getSimpleName().contains("TitlePageTab")) {
                        try {
                            Method mDate = clazz.getMethod("getProtocolDateForExcel");
                            Object val = mDate.invoke(c);
                            if (val != null) {
                                String s = val.toString().trim();
                                if (!s.isEmpty()) return s;
                            }
                        } catch (NoSuchMethodException ignored) {
                            // если геттера пока нет — молча пропускаем
                        } catch (Throwable ignored) {
                        }
                    }
                }
            }
        } catch (Throwable ignored) {
        }

        return null;
    }
    /** Создаёт (при необходимости) строки/ячейки, объединяет диапазон и задаёт текст + стиль. */
    private static void setMergedText(Sheet sheet,
                                      CellStyle style,
                                      int firstRow, int lastRow,
                                      int firstCol, int lastCol,
                                      String text) {
        // Объединяем область
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);

        // Заполняем все ячейки стилем, а текст кладём в левую верхнюю
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) cell = row.createCell(c);
                cell.setCellStyle(style);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(text);
                }
            }
        }
    }

    /** Настроить один лист: 1 страница по ширине, произвольная по высоте. */
    private static void tuneSheetToOnePageWidth(org.apache.poi.ss.usermodel.Sheet sh) {
        try {
            sh.setFitToPage(true);
            sh.setAutobreaks(true);
            sh.setHorizontallyCenter(true); // чтобы было по центру при печати

            org.apache.poi.ss.usermodel.PrintSetup ps = sh.getPrintSetup();
            ps.setPaperSize(org.apache.poi.ss.usermodel.PrintSetup.A4_PAPERSIZE);
            ps.setLandscape(true);          // обычно удобнее для широких таблиц
            ps.setFitWidth((short) 1);      // 1 страница по ширине
            ps.setFitHeight((short) 0);     // по высоте — сколько понадобится

            // Чуть ужмём поля, чтобы выиграть место по ширине
            sh.setMargin(org.apache.poi.ss.usermodel.Sheet.LeftMargin, 0.25);
            sh.setMargin(org.apache.poi.ss.usermodel.Sheet.RightMargin, 0.25);
            // Top/Bottom оставлю дефолтными; при желании можно тоже уменьшить:
            // sh.setMargin(Sheet.TopMargin, 0.5);
            // sh.setMargin(Sheet.BottomMargin, 0.5);
        } catch (Throwable ignore) {
            // не мешаем экспорту даже если где-то не поддержано
        }
    }

}
