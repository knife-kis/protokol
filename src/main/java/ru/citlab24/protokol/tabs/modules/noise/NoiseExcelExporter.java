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
                              Map<NoiseTestKind, String> dateLines,
                              Map<String, double[]> thresholds) {
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
            // ВАЖНО: передаём thresholds + kind
            NoiseLiftSheetWriter.appendLiftRoomBlocks(
                    wb, day, building, byKey, false, thresholds, NoiseTestKind.LIFT_DAY);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, night, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_NIGHT));
            NoiseLiftSheetWriter.appendLiftRoomBlocksFromRow(
                    wb, night, building, byKey, 2, true, thresholds, NoiseTestKind.LIFT_NIGHT);

            // ИТО
            NoiseSheetCommon.writeSimpleHeader(
                    wb, nonResIto,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_NONRES));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoNonResRoomBlocksFromRow(
                    wb, nonResIto, building, byKey, 2, thresholds, NoiseTestKind.ITO_NONRES);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoDay,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_DAY));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoDay, building, byKey, 2, thresholds, NoiseTestKind.ITO_RES_DAY);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_NIGHT));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoNight, building, byKey, 2, thresholds, NoiseTestKind.ITO_RES_NIGHT);

            // Авто
            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoDay, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_DAY));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoDay, building, byKey, 2, thresholds, NoiseTestKind.AUTO_DAY);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_NIGHT));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoNight, building, byKey, 2, thresholds, NoiseTestKind.AUTO_NIGHT);

            // Площадка (улица)
            NoiseSheetCommon.writeSimpleHeader(
                    wb, site, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.SITE));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSiteSheetWriter.appendSiteRoomRowsFromRow(
                    wb, site, building, byKey, 2, thresholds, NoiseTestKind.SITE);

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

    /** Совместимый враппер: если порогов не передали — считаем пустыми. */
    public static void export(Building building,
                              Map<String, DatabaseManager.NoiseValue> byKey,
                              Component parent,
                              Map<NoiseTestKind, String> dateLines) {
        export(building, byKey, parent, dateLines, java.util.Collections.emptyMap());
    }

    /** Заполняет первую строку блока (t1) случайными значениями по порогам в столбцах:
     *  T–V (Eq, пишем в T) и W–Y (MAX, пишем в W).
     *  Применяется для Лифт/ИТО/Авто/Улица в зависимости от kind.
     */
    // внутри NoiseExcelExporter
    public static void fillEqMaxFirstRow(org.apache.poi.ss.usermodel.Sheet sh,
                                         int rowIndexT1,
                                         ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind kind,
                                         Map<String, double[]> thresholdsMap) {
        String key = thresholdKeyFor(kind);
        if (thresholdsMap == null || !thresholdsMap.containsKey(key)) return;

        double[] arr = thresholdsMap.get(key);
        if (arr == null || arr.length < 4) return;

        Double eqVal  = randomBetween(arr[0], arr[1]); // Eq мин..макс
        Double maxVal = randomBetween(arr[2], arr[3]); // MAX мин..макс

        org.apache.poi.ss.usermodel.Row r = sh.getRow(rowIndexT1);
        if (r == null) r = sh.createRow(rowIndexT1);

        final int COL_T = 19; // T
        final int COL_W = 22; // W

        if (eqVal != null) {
            org.apache.poi.ss.usermodel.Cell cT = r.getCell(COL_T);
            if (cT == null) cT = r.createCell(COL_T);
            cT.setCellValue(eqVal);
        }

        if (maxVal != null) {
            org.apache.poi.ss.usermodel.Cell cW = r.getCell(COL_W);
            if (cW == null) cW = r.createCell(COL_W);
            cW.setCellValue(maxVal);
        }
    }
    /** Ключ порогов: "<источник>|<вариант>" */
    static String thresholdKeyFor(ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind kind) {
        switch (kind) {
            case LIFT_DAY:   return "Лифт|день";
            case LIFT_NIGHT: return "Лифт|ночь";
            // если у тебя есть офисная версия лифта отдельным видом — добавь сюда:
            // case LIFT_OFFICE: return "Лифт|офис";

            case ITO_NONRES:    return "ИТО|офис"; // «шум неж ИТО»
            case ITO_RES_DAY:   return "ИТО|день";
            case ITO_RES_NIGHT: return "ИТО|ночь";

            case AUTO_DAY:   return "Авто|день";
            case AUTO_NIGHT: return "Авто|ночь";

            case SITE:       return "Улица|диапазон";

            default: return "";
        }
    }
    /** Рандом в [min; max] с округлением до 1 знака. null если границы не заданы. */
    static Double randomBetween(double min, double max) {
        if (Double.isNaN(min) || Double.isNaN(max)) return null;
        if (min > max) { double t = min; min = max; max = t; }
        double v = java.util.concurrent.ThreadLocalRandom.current().nextDouble(min, Math.nextUp(max));
        return Math.round(v * 10.0) / 10.0;
    }

}
