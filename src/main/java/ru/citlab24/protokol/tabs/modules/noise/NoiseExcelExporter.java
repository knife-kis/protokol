package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.modules.noise.excel.NoiseLiftSheetWriter;
import ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

public final class NoiseExcelExporter {

    private NoiseExcelExporter() {}

    public static void export(Building building,
                              Map<String, DatabaseManager.NoiseValue> byKey,
                              Component parent,
                              java.util.Map<NoiseTestKind, String> dateLines) {
        try (Workbook wb = new XSSFWorkbook()) {

            Sheet day         = wb.createSheet("шум лифт день");
            Sheet night       = wb.createSheet("шум лифт ночь");
            Sheet nonResIto   = wb.createSheet("шум неж ИТО");
            Sheet resItoDay   = wb.createSheet("шум жил ИТО день");
            Sheet resItoNight = wb.createSheet("шум жил ИТО ночь");
            Sheet autoDay     = wb.createSheet("шум авто день");
            Sheet autoNight   = wb.createSheet("шум авто ночь");
            Sheet site        = wb.createSheet("шум площадка");

            for (Sheet sh : new Sheet[]{day, night, nonResIto, resItoDay, resItoNight, autoDay, autoNight, site}) {
                NoiseSheetCommon.setupPage(sh);
                NoiseSheetCommon.setupColumns(sh);
                NoiseSheetCommon.shrinkColumnsToOnePrintedPage(sh);
            }

            // Лифты
            NoiseLiftSheetWriter.writeLiftHeader(
                    wb, day,  NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_DAY));
            NoiseLiftSheetWriter.appendLiftRoomBlocks(wb, day, building, byKey, false);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, night, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_NIGHT));
            NoiseLiftSheetWriter.appendLiftRoomBlocksFromRow(wb, night, building, byKey, 2, true);

            // ИТО
            NoiseSheetCommon.writeSimpleHeader(
                    wb, nonResIto,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_NONRES));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter
                    .appendItoNonResRoomBlocksFromRow(wb, nonResIto, building, byKey, 2);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoDay,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_DAY));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter
                    .appendItoResRoomBlocksFromRow(wb, resItoDay, building, byKey, 2);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_NIGHT));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter
                    .appendItoResRoomBlocksFromRow(wb, resItoNight, building, byKey, 2);

            // Авто
            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoDay,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_DAY));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter
                    .appendAutoRoomBlocksFromRow(wb, autoDay, building, byKey, 2);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_NIGHT));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter
                    .appendAutoRoomBlocksFromRow(wb, autoNight, building, byKey, 2);

            // Площадка (улица) — НОВОЕ: один ряд на т1/т2/т3; C — только комната; D — авто/поезд; T–V и W–Y не объединяем (пусто)
            NoiseSheetCommon.writeSimpleHeader(
                    wb, site, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.SITE));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSiteSheetWriter
                    .appendSiteRoomRowsFromRow(wb, site, building, byKey, 2);

            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Сохранить Excel (шумы)");
            fc.setSelectedFile(new File("шумы.xlsx"));
            if (fc.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = NoiseSheetCommon.ensureXlsx(fc.getSelectedFile());
                try (FileOutputStream out = new FileOutputStream(file)) {
                    wb.write(out);
                }
                JOptionPane.showMessageDialog(parent, "Файл сохранён:\n" + file.getAbsolutePath(),
                        "Экспорт", JOptionPane.INFORMATION_MESSAGE);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent, "Ошибка экспорта: " + ex.getMessage(),
                    "Экспорт", JOptionPane.ERROR_MESSAGE);
        }
    }
}
