package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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

            // >>> НОВОЕ: привести ВСЕ листы к печати "в 1 страницу по ширине"
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
