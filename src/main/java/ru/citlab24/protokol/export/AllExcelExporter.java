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

    public static void exportAll(MainFrame frame, Building building, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = new XSSFWorkbook()) {
            // 1) Микроклимат (есть append)
            MicroclimateExcelExporter.appendToWorkbook(building, -1, wb);

            // 2) Вентиляция (берём текущие записи из вкладки)
            List<VentilationRecord> ventRecords = null;
            if (frame != null && frame.getVentilationTab() != null) {
                ventRecords = frame.getVentilationTab().getRecordsForExport();
            }
            if (ventRecords != null && !ventRecords.isEmpty()) {
                VentilationExcelExporter.appendToWorkbook(ventRecords, wb);
            }

            // 3) Освещение — пробуем вызвать LightingExcelExporter.appendToWorkbook(building, int, wb)
            tryInvokeAppend("ru.citlab24.protokol.tabs.modules.lighting.LightingExcelExporter",
                    new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                    new Object[]{building, -1, wb});

            // 4) Радиация — пробуем вызвать RadiationExcelExporter.appendToWorkbook(building, int, wb)
            tryInvokeAppend("ru.citlab24.protokol.tabs.modules.med.RadiationExcelExporter",
                    new Class[]{ru.citlab24.protokol.tabs.models.Building.class, int.class, Workbook.class},
                    new Object[]{building, -1, wb});

            // 5) Один общий диалог сохранения
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
            ex.printStackTrace();
            JOptionPane.showMessageDialog(parent,
                    "Ошибка объединённого экспорта: " + ex.getMessage(),
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
}
