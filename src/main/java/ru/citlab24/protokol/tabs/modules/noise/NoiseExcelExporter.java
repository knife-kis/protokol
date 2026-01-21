package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.CellStyle;
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
    private static final java.util.Map<Workbook, java.util.Map<CellStyle, CellStyle>> ONE_DECIMAL_STYLES =
            new java.util.WeakHashMap<>();
    private static final java.util.Map<RangeKey, java.util.Deque<Double>> RECENT_BY_RANGE =
            new java.util.HashMap<>();
    public static void export(Building building,
                              Map<String, DatabaseManager.NoiseValue> byKey,
                              Component parent,
                              Map<NoiseTestKind, String> dateLines,
                              Map<String, double[]> thresholds) {
        try (Workbook wb = new XSSFWorkbook()) {
            RECENT_BY_RANGE.clear();

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

            int nextNo = 1;

            // Лифты
            NoiseLiftSheetWriter.writeLiftHeader(
                    wb, day,  NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_DAY));
            // ВАЖНО: передаём thresholds + kind
            nextNo = NoiseLiftSheetWriter.appendLiftRoomBlocks(
                    wb, day, building, byKey, false, nextNo, thresholds, NoiseTestKind.LIFT_DAY);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, night, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_NIGHT));
            nextNo = NoiseLiftSheetWriter.appendLiftRoomBlocksFromRow(
                    wb, night, building, byKey, 2, nextNo, true, thresholds, NoiseTestKind.LIFT_NIGHT);

            // ИТО
            NoiseSheetCommon.writeSimpleHeader(
                    wb, nonResIto,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_NONRES));
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoNonResRoomBlocksFromRow(
                    wb, nonResIto, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_NONRES, true);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoDay,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_DAY));
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoDay, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_RES_DAY, false);
            int zumStartRow = Math.max(2, resItoDay.getLastRowNum() + 1);
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendZumResRoomBlocksFromRow(
                    wb, resItoDay, building, byKey, zumStartRow, nextNo, thresholds, NoiseTestKind.ZUM_DAY, true);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_NIGHT));
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoNight, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_RES_NIGHT, true);

            // Авто
            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoDay, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_DAY));
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoDay, building, byKey, 2, nextNo, thresholds, NoiseTestKind.AUTO_DAY);

            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_NIGHT));
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoNight, building, byKey, 2, nextNo, thresholds, NoiseTestKind.AUTO_NIGHT);

            // Площадка (улица)
            NoiseSheetCommon.writeSimpleHeader(
                    wb, site, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.SITE));
            ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSiteSheetWriter.appendSiteRoomRowsFromRow(
                    wb, site, building, byKey, 2, nextNo, thresholds, NoiseTestKind.SITE);

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
            cT.setCellStyle(oneDecimalStyle(sh.getWorkbook(), cT.getCellStyle()));
        }

        if (maxVal != null) {
            org.apache.poi.ss.usermodel.Cell cW = r.getCell(COL_W);
            if (cW == null) cW = r.createCell(COL_W);
            cW.setCellValue(maxVal);
            cW.setCellStyle(oneDecimalStyle(sh.getWorkbook(), cW.getCellStyle()));
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
            case ZUM_DAY:       return "Зум|день";

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
        java.util.List<Double> candidates = buildCandidates(min, max);
        if (candidates.isEmpty()) return null;

        RangeKey key = new RangeKey(min, max);
        java.util.Deque<Double> recent = RECENT_BY_RANGE.computeIfAbsent(key, k -> new java.util.ArrayDeque<>());
        java.util.Set<Double> recentSet = new java.util.HashSet<>(recent);

        java.util.List<Double> available = new java.util.ArrayList<>(candidates.size());
        for (Double value : candidates) {
            if (!recentSet.contains(value)) {
                available.add(value);
            }
        }
        java.util.List<Double> source = available.isEmpty() ? candidates : available;
        int pick = java.util.concurrent.ThreadLocalRandom.current().nextInt(source.size());
        Double chosen = source.get(pick);

        recent.addLast(chosen);
        while (recent.size() > 6) {
            recent.removeFirst();
        }
        return chosen;
    }

    private static java.util.List<Double> buildCandidates(double min, double max) {
        double start = Math.round(min * 10.0) / 10.0;
        double end = Math.round(max * 10.0) / 10.0;
        if (start > end) { double t = start; start = end; end = t; }
        java.util.List<Double> values = new java.util.ArrayList<>();
        for (double v = start; v <= end + 0.0000001; v += 0.1) {
            values.add(Math.round(v * 10.0) / 10.0);
        }
        return values;
    }

    private static CellStyle oneDecimalStyle(Workbook wb, CellStyle baseStyle) {
        java.util.Map<CellStyle, CellStyle> cache = ONE_DECIMAL_STYLES.computeIfAbsent(
                wb, key -> new java.util.IdentityHashMap<>());
        return cache.computeIfAbsent(baseStyle, style -> {
            CellStyle oneDecimal = wb.createCellStyle();
            oneDecimal.cloneStyleFrom(style);
            oneDecimal.setDataFormat(wb.createDataFormat().getFormat("0.0"));
            return oneDecimal;
        });
    }

    private static final class RangeKey {
        private final long minDec;
        private final long maxDec;

        private RangeKey(double min, double max) {
            this.minDec = Math.round(min * 10.0);
            this.maxDec = Math.round(max * 10.0);
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof RangeKey)) return false;
            RangeKey rangeKey = (RangeKey) o;
            return minDec == rangeKey.minDec && maxDec == rangeKey.maxDec;
        }

        @Override
        public int hashCode() {
            return java.util.Objects.hash(minDec, maxDec);
        }
    }

}
