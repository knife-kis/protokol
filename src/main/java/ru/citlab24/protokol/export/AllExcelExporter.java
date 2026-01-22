package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.modules.microclimateTab.MicroclimateExcelExporter;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationExcelExporter;
import ru.citlab24.protokol.tabs.modules.ventilation.VentilationRecord;
import ru.citlab24.protokol.tabs.titleTab.TitlePageTab;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Method;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.List;
import java.util.Locale;


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
            if (hasAnyMicroclimateSelections(building)) {
                MicroclimateExcelExporter.appendToWorkbook(
                        building,
                        -1,
                        wb,
                        MicroclimateExcelExporter.TemperatureMode.COLD
                );
            }

            // 2) Вентиляция (берём текущие записи из вкладки)
            List<VentilationRecord> ventRecords = null;
            if (frame != null && frame.getVentilationTab() != null) {
                ventRecords = frame.getVentilationTab().getRecordsForExport();
            }
            if (ventRecords != null && !ventRecords.isEmpty()) {
                VentilationExcelExporter.appendToWorkbook(ventRecords, wb);
            }

            // 3) Радиация
            if (hasAnyRadiationSelections(building)) {
                tryInvokeAppend("ru.citlab24.protokol.tabs.modules.med.RadiationExcelExporter",
                        new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                        new Object[]{building, -1, wb});
            }

            // 4) Искусственное освещение
            java.util.Map<Integer, Boolean> litMap = null;
            try {
                if (frame != null && frame.getArtificialLightingTab() != null) {
                    litMap = frame.getArtificialLightingTab().snapshotSelectionMap();
                }
            } catch (Throwable ignore) {}
            if (hasAnyArtificialLightingSelections(litMap)) {
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
            }

            // 5) «Иск освещение (2)» — ближе к концу
            appendStreetLightingSheet(wb, building);

            // 6) КЕО — последним листом
            if (hasAnyLightingSelections(building)) {
                tryInvokeAppend("ru.citlab24.protokol.tabs.modules.lighting.LightingExcelExporter",
                        new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                        new Object[]{building, -1, wb});
            }

            // Обновляем строку с перечнем показателей на титульной странице
            IndicatorsTextUpdater.updateIndicatorsText(wb);

// ПЕРЕСТРАИВАЕМ таблицу "Сведения о средствах измерения" уже ПОСЛЕ того,
// как появились листы и обновился перечень показателей
            TitlePageMeasurementTableWriter.rebuildTable(wb, buildAdditionalInfoText(frame));

// Приводим ВСЕ листы к печати "в 1 страницу по ширине"
            applyFitToPageWidthForAllSheets(wb);
            PrintSetupUtils.applyDuplexShortEdge(wb);

            // Общий колонтитул для всех листов, кроме титульного
            applyCommonHeaderToSheets(wb, frame);

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
        PrintSetupUtils.applyDuplexShortEdge(sheet);

        // Базовый стиль: Arial 10, перенос по словам
        Font baseFont = wb.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 10);

        Font boldFont = wb.createFont();
        boldFont.setFontName("Arial");
        boldFont.setFontHeightInPoints((short) 10);
        boldFont.setBold(true);

        Font smallFont = wb.createFont();
        smallFont.setFontName("Arial");
        smallFont.setFontHeightInPoints((short) 8);

        CellStyle baseStyle = wb.createCellStyle();
        baseStyle.setFont(baseFont);
        baseStyle.setWrapText(true);
        baseStyle.setVerticalAlignment(VerticalAlignment.TOP);
        baseStyle.setAlignment(HorizontalAlignment.LEFT);

        CellStyle centerMiddleStyle = wb.createCellStyle();
        centerMiddleStyle.cloneStyleFrom(baseStyle);
        centerMiddleStyle.setAlignment(HorizontalAlignment.CENTER);
        centerMiddleStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle boldCenterStyle = wb.createCellStyle();
        boldCenterStyle.cloneStyleFrom(centerMiddleStyle);
        boldCenterStyle.setFont(boldFont);

        CellStyle headerBorderStyle = wb.createCellStyle();
        headerBorderStyle.cloneStyleFrom(centerMiddleStyle);
        setThinBorders(headerBorderStyle);

        CellStyle measurementStyle = wb.createCellStyle();
        measurementStyle.cloneStyleFrom(centerMiddleStyle);
        measurementStyle.setFont(smallFont);
        setThinBorders(measurementStyle);

        // Высоты строк: фиксированный список (значения заданы в пикселях).
        float[] rowHeightsPx = new float[] {
                20f, 21f, 10f, 20f, 16f, 20f, 20f, 47f, 20f, 20f,
                19f, 20f, 20f, 20f, 20f, 20f, 9f, 20f, 9f, 20f,
                9f, 20f, 9f, 20f, 9f, 20f, 9f, 20f, 9f, 20f,
                9f, 20f, 9f, 20f, 9f, 20f, 126f
        };// Высоты строк оставляем как раньше — тебе они подошли
        for (int r = 0; r < rowHeightsPx.length; r++) {
            Row row = sheet.getRow(r);
            if (row == null) row = sheet.createRow(r);
            row.setHeightInPoints(pixelsToPoints(rowHeightsPx[r]));
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
        setMergedText(sheet, boldCenterStyle, 0, 2, 0, 8, txtA1);

        // 2) S1-Z3
        String txtS1 = "УТВЕРЖДАЮ\nЗаместитель заведующего лабораторией";
        setMergedText(sheet, boldCenterStyle, 0, 2, 18, 25, txtS1);

        // 3) A4-I5
        String txtA4 = "660064, Красноярский край, г.о. Красноярск,\n" +
                "г. Красноярск, ул. Парусная, д. 8, кв. 99";
        setMergedText(sheet, centerMiddleStyle, 3, 4, 0, 8, txtA4);

        // 4) A6-I8
        String txtA6 = "Испытательная лаборатория\n" +
                "Общества с ограниченной ответственностью\n" +
                "«Центр исследовательских технологий»\n" +
                "Уникальный номер записи в Реестре аккредитованных лиц RA.RU.21ОХ37 от 07.04.2023";
        XSSFRichTextString richA6 = new XSSFRichTextString(txtA6);
        int registryIndex = txtA6.lastIndexOf("Уникальный номер записи");
        int boldEnd = (registryIndex >= 0) ? registryIndex : txtA6.length();
        richA6.applyFont(0, boldEnd, boldFont);
        if (boldEnd < txtA6.length()) {
            richA6.applyFont(boldEnd, txtA6.length(), baseFont);
        }
        setMergedRichText(sheet, centerMiddleStyle, 5, 7, 0, 8, richA6);

        // 5) A9-I11
        String txtA9 = "660041, Красноярский край, г. Красноярск,\n" +
                "пр. Свободный дом 75, пом. 9 кабинет 12\n" +
                "+7 (391) 244-03-10, cit-krsk@yandex.ru";
        setMergedText(sheet, centerMiddleStyle, 8, 10, 0, 8, txtA9);

        // S5-Z5: _______________/М.Е. Гаврилова/
        String txtS5 = "_______________/М.Е. Гаврилова/";
        setMergedText(sheet, boldCenterStyle, 4, 4, 18, 25, txtS5);

        // S7-Z7: дата из программы в формате "12 декабря 2025 г."
        String protocolDate = formatProtocolDateForHeader(resolveProtocolDateText(frame));
        setMergedText(sheet, boldCenterStyle, 6, 6, 18, 25, protocolDate);

        // S9-Z9: МП
        setMergedText(sheet, boldCenterStyle, 8, 8, 18, 25, "МП");

        TitlePageValues titleValues = readTitlePageValues(frame);

        // A13-Z13 — заголовок с номером заявки и суффиксом Ф/XX
        String protocolTitle = buildProtocolTitle(titleValues.applicationNumber);
        setMergedText(sheet, boldCenterStyle, 12, 12, 0, 25, protocolTitle);

        // A14-Z14 — Вид испытаний
        setMergedText(sheet, centerMiddleStyle, 13, 13, 0, 25,
                "Вид испытаний: измерения физических факторов");

        // Табличные строки 1–6
        setCellValue(sheet, centerMiddleStyle, 15, 0, "1.");
        setMergedText(sheet, baseStyle, 15, 15, 1, 25,
                "Наименование и контактные данные заявителя (заказчика): " +
                        safe(titleValues.customerNameAndContacts));

        setCellValue(sheet, centerMiddleStyle, 17, 0, "2.");
        setMergedText(sheet, baseStyle, 17, 17, 1, 25,
                "Юридический адрес заказчика: " + safe(titleValues.customerLegalAddress));

        setCellValue(sheet, centerMiddleStyle, 19, 0, "3.");
        setMergedText(sheet, baseStyle, 19, 19, 1, 25,
                "Фактический адрес заказчика: " + safe(titleValues.customerActualAddress));

        setCellValue(sheet, centerMiddleStyle, 21, 0, "4.");
        String objectNameText = "Наименование предприятия, организации, объекта, где производились измерения: " +
                safe(titleValues.objectName);
        setMergedText(sheet, baseStyle, 21, 21, 1, 25, objectNameText);
        adjustRowHeightForMergedTextDoubling(sheet, 21, 1, 25, objectNameText);

        setCellValue(sheet, centerMiddleStyle, 23, 0, "5.");
        setMergedText(sheet, baseStyle, 23, 23, 1, 25,
                "Адрес предприятия (объекта): " + safe(titleValues.objectAddress));

        setCellValue(sheet, centerMiddleStyle, 25, 0, "6.");
        setMergedText(sheet, baseStyle, 25, 25, 1, 25, buildBasisLine(titleValues));

        setCellValue(sheet, centerMiddleStyle, 27, 0, "7.");
        setMergedText(sheet, baseStyle, 27, 27, 1, 25,
                "Измерения проводились в присутствии представителя заказчика: " +
                        safe(titleValues.representative));

        setCellValue(sheet, centerMiddleStyle, 29, 0, "8.");
        String indicatorsText = "Показатели, по которым проводились измерения:";
        setMergedText(sheet, baseStyle, 29, 29, 1, 25, indicatorsText);
        adjustRowHeightForMergedText(sheet, 29, 1, 25, indicatorsText);

        setCellValue(sheet, centerMiddleStyle, 31, 0, "9.");
        setMergedText(sheet, baseStyle, 31, 31, 1, 25,
                "Регистрационный номер карты замеров: " + buildMeasurementCardNumber(titleValues.applicationNumber));

        setCellValue(sheet, centerMiddleStyle, 33, 0, "10.");
        setMergedText(sheet, baseStyle, 33, 33, 1, 25,
                "Сведения о средствах измерения:");

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
    // === НОВОЕ: «Иск освещение (2)» — собираем строки и добавляем лист в текущую книгу ===
    private static void appendStreetLightingSheet(Workbook wb, Building building) {
        try {
            // 1) Подтянуть сохранённые 4 значения по ключам (если есть)
            java.util.Map<String, Double[]> byKey = java.util.Collections.emptyMap();
            try {
                byKey = ru.citlab24.protokol.db.DatabaseManager
                        .loadStreetLightingValuesByKey(building.getId());
            } catch (java.sql.SQLException ex) {
                System.err.println("[AllExcelExporter] WARN: не удалось прочитать 'Иск освещение (2)' из БД: " + ex.getMessage());
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

                        // Ключ как во вкладке StreetLightingTab/DatabaseManager:
                        // ID|roomId (если есть) иначе sectionIndex|этаж|помещение|комната
                        String key = buildStreetLightingKey(f, s, r, floorPart, spacePart, roomPart);

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

            if (!hasAnyStreetLightingValues(rows)) {
                return;
            }

            // 3) Вставить лист «Иск освещение (2)» в текущую книгу
            ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.appendToWorkbook(rows, wb);

            // 4) Попробовать зафиксировать порядок: «Иск освещение (2)» перед «КЕО»
            try {
                int streetIdx = wb.getSheetIndex("Иск освещение (2)");
                int keoIdx    = wb.getSheetIndex("КЕО");
                if (streetIdx >= 0 && keoIdx >= 0 && streetIdx > keoIdx) {
                    wb.setSheetOrder("Иск освещение (2)", keoIdx);
                }
            } catch (Throwable ignore) {
                // на случай, если лист КЕО назван иначе — порядок уже обеспечен самим вызовом до КЕО
            }
        } catch (Throwable t) {
            System.err.println("[AllExcelExporter] Ошибка добавления 'Иск освещение (2)': " + t.getMessage());
        }
    }
    private static String buildStreetLightingKey(ru.citlab24.protokol.tabs.models.Floor f,
                                                 ru.citlab24.protokol.tabs.models.Space s,
                                                 ru.citlab24.protokol.tabs.models.Room r,
                                                 String floorPart,
                                                 String spacePart,
                                                 String roomPart) {
        Integer stableId = resolveStableRoomId(r);
        if (stableId != null) {
            return "ID|" + stableId;
        }
        int sectionIndex = (f != null) ? f.getSectionIndex() : 0;
        return sectionIndex + "|" + floorPart + "|" + spacePart + "|" + roomPart;
    }
    private static Integer resolveStableRoomId(ru.citlab24.protokol.tabs.models.Room r) {
        if (r == null) return null;
        Integer original = r.getOriginalRoomId();
        if (original != null && original > 0) return original;
        int id = r.getId();
        return (id > 0) ? id : null;
    }

    private static boolean hasAnyMicroclimateSelections(Building building) {
        if (building == null || building.getFloors() == null) return false;
        for (ru.citlab24.protokol.tabs.models.Floor floor : building.getFloors()) {
            if (floor == null || floor.getSpaces() == null) continue;
            for (ru.citlab24.protokol.tabs.models.Space space : floor.getSpaces()) {
                if (space == null || space.getRooms() == null) continue;
                for (ru.citlab24.protokol.tabs.models.Room room : space.getRooms()) {
                    if (room != null && room.isMicroclimateSelected()) return true;
                }
            }
        }
        return false;
    }

    private static boolean hasAnyRadiationSelections(Building building) {
        if (building == null || building.getFloors() == null) return false;
        for (ru.citlab24.protokol.tabs.models.Floor floor : building.getFloors()) {
            if (floor == null || floor.getSpaces() == null) continue;
            for (ru.citlab24.protokol.tabs.models.Space space : floor.getSpaces()) {
                if (space == null || space.getRooms() == null) continue;
                for (ru.citlab24.protokol.tabs.models.Room room : space.getRooms()) {
                    if (room != null && room.isRadiationSelected()) return true;
                }
            }
        }
        return false;
    }

    private static boolean hasAnyLightingSelections(Building building) {
        if (building == null || building.getFloors() == null) return false;
        for (ru.citlab24.protokol.tabs.models.Floor floor : building.getFloors()) {
            if (floor == null || floor.getSpaces() == null) continue;
            for (ru.citlab24.protokol.tabs.models.Space space : floor.getSpaces()) {
                if (space == null || space.getRooms() == null) continue;
                for (ru.citlab24.protokol.tabs.models.Room room : space.getRooms()) {
                    if (room != null && room.isSelected()) return true;
                }
            }
        }
        return false;
    }

    private static boolean hasAnyArtificialLightingSelections(java.util.Map<Integer, Boolean> selectionMap) {
        if (selectionMap == null || selectionMap.isEmpty()) return false;
        for (Boolean value : selectionMap.values()) {
            if (Boolean.TRUE.equals(value)) return true;
        }
        return false;
    }

    private static boolean hasAnyStreetLightingValues(
            java.util.List<ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.RowData> rows) {
        if (rows == null || rows.isEmpty()) return false;
        for (ru.citlab24.protokol.tabs.modules.lighting.StreetLightingExcelExporter.RowData row : rows) {
            if (row == null) continue;
            if (row.leftMax != null || row.centerMin != null || row.rightMax != null || row.bottomMin != null) {
                return true;
            }
        }
        return false;
    }
    /** Применяет настройку печати "уместить по ширине в 1 страницу" ко всем листам книги. */
    private static void applyFitToPageWidthForAllSheets(Workbook wb) {
        if (wb == null) return;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            org.apache.poi.ss.usermodel.Sheet sh = wb.getSheetAt(i);
            if (sh != null) tuneSheetToOnePageWidth(sh);
        }
    }
    private static TitlePageValues readTitlePageValues(MainFrame frame) {
        TitlePageValues values = new TitlePageValues();
        TitlePageTab tab = findTitlePageTab(frame);
        if (tab != null) {
            values.customerNameAndContacts = safe(tab.getCustomerNameAndContacts());
            values.customerLegalAddress = safe(tab.getCustomerLegalAddress());
            values.customerActualAddress = safe(tab.getCustomerActualAddress());
            values.objectName = safe(tab.getObjectName());
            values.objectAddress = safe(tab.getObjectAddress());
            values.contractNumber = safe(tab.getContractNumber());
            values.contractDate = safe(tab.getContractDate());
            values.applicationNumber = safe(tab.getApplicationNumber());
            values.applicationDate = safe(tab.getApplicationDate());
            values.representative = safe(tab.getRepresentative());
        }
        return values;
    }

    private static String buildAdditionalInfoText(MainFrame frame) {
        TitlePageTab tab = findTitlePageTab(frame);
        List<TitlePageTab.MeasurementRowData> rows =
                (tab != null) ? tab.getMeasurementRows() : java.util.Collections.emptyList();

        String datesText = buildMeasurementDatesText(rows);
        String insideTempsText = buildInsideTemperatureText(rows);
        String outsideTempsText = buildOutsideTemperatureText(rows);

        return "12. Дополнительные сведения (характеристика объекта): " +
                "Измерения проведены в жилых, общественных помещениях и придомовой территории. " +
                "Измерения были проведены " + datesText + ". " +
                "Температура воздуха в помещениях: " + insideTempsText + ". " +
                "Измерения микроклимата проведены в зимний период. " +
                "Измерения ионизирующих излучений проведены в летний период. " +
                "Температура наружного воздуха: " + outsideTempsText + ". " +
                "Абсолютная неопределенность далее \"ΔН\" приведена в результатах измерений. " +
                "Поверхностных радиационных аномалий в конструкциях здания не обнаружено. " +
                "Перед измерением ЭРОА радона здание выдержанно в течении более 12 часов при закрытых дверях и окнах. " +
                "Измерения наружной освещенности проведены в светлое время суток при сплошной равномерной облачности, " +
                "покрывающей весь небосвод. " +
                "Измерения освещенности внутри помещения выполнены без наличия мебели, вымытых и исправных " +
                "светопрозрачных заполнениях в светопроемах, при выключенных искусственных источниках световой среды. " +
                "Измерения искусственного освещения проведены в темное время суток с отсутствием " +
                "естественного освещения. Отношение естественной освещенности к искусственной составляет менее 0,1.";
    }

    private static String buildMeasurementDatesText(List<TitlePageTab.MeasurementRowData> rows) {
        if (rows == null || rows.isEmpty()) {
            return "дд.мм.гггг";
        }
        java.util.List<String> dates = new java.util.ArrayList<>();
        for (TitlePageTab.MeasurementRowData row : rows) {
            if (row == null) continue;
            String date = safe(row.date);
            dates.add(date.isEmpty() ? "дд.мм.гггг" : date);
        }
        return String.join(", ", dates);
    }

    private static String buildInsideTemperatureText(List<TitlePageTab.MeasurementRowData> rows) {
        return buildTemperatureText(rows, true);
    }

    private static String buildOutsideTemperatureText(List<TitlePageTab.MeasurementRowData> rows) {
        return buildTemperatureText(rows, false);
    }

    private static String buildTemperatureText(List<TitlePageTab.MeasurementRowData> rows, boolean inside) {
        if (rows == null || rows.isEmpty()) {
            String temp = inside ? "___ °С до ___ °С" : "___°С до ___ °С";
            return "дд.мм.гггг составила от " + temp;
        }
        java.util.List<String> parts = new java.util.ArrayList<>();
        for (TitlePageTab.MeasurementRowData row : rows) {
            if (row == null) continue;
            String date = safe(row.date);
            if (date.isEmpty()) {
                date = "дд.мм.гггг";
            }
            String start = inside ? safe(row.tempInsideStart) : safe(row.tempOutsideStart);
            String end = inside ? safe(row.tempInsideEnd) : safe(row.tempOutsideEnd);
            if (start.isEmpty()) start = "___";
            if (end.isEmpty()) end = "___";
            if (inside) {
                parts.add(date + " составила от " + start + " °С до " + end + " °С");
            } else {
                parts.add(date + " составляла от " + start + "°С до " + end + " °С");
            }
        }
        return String.join("; ", parts);
    }

    private static String buildProtocolTitle(String applicationNumber) {
        return "Протокол испытаний № " + buildProtocolNumber(applicationNumber);
    }

    private static String buildProtocolNumber(String applicationNumber) {
        String appNumber = safe(applicationNumber);
        if (appNumber.isEmpty()) {
            appNumber = "_____";
        }
        String suffix = lastDigitsOrPlaceholder(applicationNumber, 2, "__");
        return appNumber + "-1-Ф/" + suffix;
    }

    private static String buildBasisLine(TitlePageValues values) {
        String contractNum = valueOrPlaceholder(values.contractNumber);
        String contractDate = valueOrPlaceholder(values.contractDate);
        String applicationNum = valueOrPlaceholder(values.applicationNumber);
        String applicationDate = valueOrPlaceholder(values.applicationDate);
        return "Основание для измерений: договор №" + contractNum +
                " от " + contractDate +
                ", заявка №" + applicationNum +
                " от " + applicationDate + ".";
    }

    private static String buildMeasurementCardNumber(String applicationNumber) {
        String base = valueOrPlaceholder(applicationNumber);
        return base + "-1К";
    }

    private static void setCellValue(Sheet sheet, CellStyle style, int rowIdx, int colIdx, String text) {
        Row row = sheet.getRow(rowIdx);
        if (row == null) {
            row = sheet.createRow(rowIdx);
        }
        Cell cell = row.getCell(colIdx);
        if (cell == null) {
            cell = row.createCell(colIdx);
        }
        cell.setCellStyle(style);
        cell.setCellValue(text);
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
        // 1.5) Прямой доступ к TitlePageTab, если он есть
        TitlePageTab tab = findTitlePageTab(frame);
        if (tab != null) {
            String protocolDate = safe(tab.getProtocolDate());
            if (!protocolDate.isEmpty()) {
                return protocolDate;
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

    private static TitlePageTab findTitlePageTab(MainFrame frame) {
        if (frame == null) return null;
        try {
            JTabbedPane tabs = frame.getTabbedPane();
            if (tabs == null) return null;
            for (int i = 0; i < tabs.getTabCount(); i++) {
                Component c = tabs.getComponentAt(i);
                if (c instanceof TitlePageTab) {
                    return (TitlePageTab) c;
                }
            }
        } catch (Throwable ignore) {
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

    private static void setMergedRichText(Sheet sheet,
                                          CellStyle style,
                                          int firstRow, int lastRow,
                                          int firstCol, int lastCol,
                                          XSSFRichTextString text) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sheet.addMergedRegion(region);

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

    private static void setThinBorders(CellStyle style) {
        if (style == null) return;
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }
    static void adjustRowHeightForMergedText(Sheet sheet,
                                             int rowIndex,
                                             int firstCol,
                                             int lastCol,
                                             String text) {
        if (sheet == null) return;

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double totalChars = 0.0;
        for (int c = firstCol; c <= lastCol; c++) {
            totalChars += sheet.getColumnWidth(c) / 256.0;
        }
        totalChars = Math.max(1.0, totalChars);

        int lines = estimateWrappedLines(text, totalChars);
        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        float newHeightPx = baseHeightPx * Math.max(1, lines);
        row.setHeightInPoints(pixelsToPoints(newHeightPx));
    }

    private static void adjustRowHeightForMergedTextDoubling(Sheet sheet,
                                                             int rowIndex,
                                                             int firstCol,
                                                             int lastCol,
                                                             String text) {
        if (sheet == null) return;

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        double totalChars = totalColumnChars(sheet, firstCol, lastCol);
        int lines = estimateWrappedLines(text, totalChars);

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        if (baseHeightPx <= 0f) {
            baseHeightPx = pointsToPixels(sheet.getDefaultRowHeightInPoints());
        }

        int multiplier = 1;
        while (multiplier < lines) {
            multiplier *= 2;
        }

        row.setHeightInPoints(pixelsToPoints(baseHeightPx * multiplier));
    }

    private static void adjustRowHeightForMergedSections(Sheet sheet,
                                                         int rowIndex,
                                                         int[][] ranges,
                                                         String[] texts) {
        if (sheet == null || ranges == null || texts == null || ranges.length != texts.length) {
            return;
        }

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        float baseHeightPx = pointsToPixels(row.getHeightInPoints());
        int maxLines = 1;
        for (int i = 0; i < ranges.length; i++) {
            int[] range = ranges[i];
            if (range == null || range.length != 2) continue;
            double totalChars = totalColumnChars(sheet, range[0], range[1]);
            int lines = estimateWrappedLines(texts[i], totalChars);
            maxLines = Math.max(maxLines, lines);
        }

        row.setHeightInPoints(pixelsToPoints(baseHeightPx * maxLines));
    }

    private static double totalColumnChars(Sheet sheet, int firstCol, int lastCol) {
        double totalChars = 0.0;
        for (int c = firstCol; c <= lastCol; c++) {
            totalChars += sheet.getColumnWidth(c) / 256.0;
        }
        return Math.max(1.0, totalChars);
    }

    private static int estimateWrappedLines(String text, double colChars) {
        if (text == null || text.isBlank()) return 1;

        int lines = 0;
        String[] segments = text.split("\\r?\\n");
        for (String seg : segments) {
            int len = Math.max(1, seg.trim().length());
            lines += (int) Math.ceil(len / Math.max(1.0, colChars));
        }
        return Math.max(1, lines);
    }

    private static String safe(String value) {
        return (value == null) ? "" : value.trim();
    }

    private static void applyCommonHeaderToSheets(Workbook wb, MainFrame frame) {
        if (wb == null || wb.getNumberOfSheets() <= 1) {
            return;
        }

        String dateText = formatProtocolDateForHeader(resolveProtocolDateText(frame));
        String protocolNumber = resolveProtocolNumber(frame);

        String headerText = "&\"Arial\"&9" +
                "Протокол от " + dateText +
                " № " + protocolNumber +
                "\n" +
                "Общее количество страниц &[Страниц] Страница &[Страница]";

        for (int i = 1; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) continue;
            Footer footer = sheet.getFooter();
            if (footer != null) {
                footer.setRight(headerText);
            }
        }
    }

    private static String resolveProtocolNumber(MainFrame frame) {
        TitlePageTab tab = findTitlePageTab(frame);
        if (tab != null) {
            String appNumber = safe(tab.getApplicationNumber());
            if (!appNumber.isEmpty()) {
                return buildProtocolNumber(appNumber);
            }
        }
        return buildProtocolNumber("");
    }

    private static String formatProtocolDateForHeader(String rawDate) {
        String trimmed = safe(rawDate);
        if (trimmed.isEmpty()) {
            return "___ __________ ____ г.";
        }

        DateTimeFormatter parser = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        DateTimeFormatter printer = DateTimeFormatter.ofPattern("d MMMM yyyy", new Locale("ru"));
        try {
            LocalDate date = LocalDate.parse(trimmed, parser);
            return printer.format(date) + " г.";
        } catch (DateTimeParseException ignore) {
        }

        if (trimmed.endsWith("г.") || trimmed.endsWith("г")) {
            return trimmed;
        }
        return trimmed + " г.";
    }

    private static String valueOrPlaceholder(String value) {
        String trimmed = safe(value);
        return trimmed.isEmpty() ? "_____" : trimmed;
    }

    private static String lastDigitsOrPlaceholder(String text, int count, String placeholder) {
        if (text == null) {
            return placeholder;
        }
        StringBuilder digitsOnly = new StringBuilder();
        for (char ch : text.toCharArray()) {
            if (Character.isDigit(ch)) {
                digitsOnly.append(ch);
            }
        }
        if (digitsOnly.length() == 0) {
            return placeholder;
        }
        if (digitsOnly.length() <= count) {
            return digitsOnly.toString();
        }
        return digitsOnly.substring(digitsOnly.length() - count);
    }

    private static final class TitlePageValues {
        String customerNameAndContacts = "";
        String customerLegalAddress = "";
        String customerActualAddress = "";
        String objectName = "";
        String objectAddress = "";
        String contractNumber = "";
        String contractDate = "";
        String applicationNumber = "";
        String applicationDate = "";
        String representative = "";
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
            PrintSetupUtils.applyDuplexShortEdge(sh);

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
    private static float pixelsToPoints(float pixels) {
        return pixels * 72f / 96f;
    }

    private static float pointsToPixels(float points) {
        return points * 96f / 72f;
    }
}
