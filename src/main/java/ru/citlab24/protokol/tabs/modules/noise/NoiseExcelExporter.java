package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTHeaderFooter;
import ru.citlab24.protokol.MainFrame;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.export.AllExcelExporter;
import ru.citlab24.protokol.tabs.models.Building;
import ru.citlab24.protokol.tabs.modules.noise.excel.NoiseLiftSheetWriter;
import ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;
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

            AllExcelExporter.appendNoiseTitleSheet(findMainFrame(parent), building, wb);

            Sheet day         = wb.createSheet("шум лифт день");
            Sheet night       = wb.createSheet("шум лифт ночь");
            Sheet nonResIto   = wb.createSheet("шум неж ИТО");
            Sheet resItoDay   = wb.createSheet("шум жил ИТО день");
            Sheet resItoNight = wb.createSheet("шум жил ИТО ночь");
            Sheet autoDay     = wb.createSheet("шум авто день");
            Sheet autoNight   = wb.createSheet("шум авто ночь");
            Sheet site        = wb.createSheet("шум площадка");
            List<SheetExportState> sheetStates = new ArrayList<>();

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
            int beforeNo = nextNo;
            nextNo = NoiseLiftSheetWriter.appendLiftRoomBlocks(
                    wb, day, building, byKey, false, nextNo, thresholds, NoiseTestKind.LIFT_DAY);
            sheetStates.add(new SheetExportState(day, nextNo > beforeNo));

            NoiseSheetCommon.writeSimpleHeader(
                    wb, night, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.LIFT_NIGHT));
            beforeNo = nextNo;
            nextNo = NoiseLiftSheetWriter.appendLiftRoomBlocksFromRow(
                    wb, night, building, byKey, 2, nextNo, true, thresholds, NoiseTestKind.LIFT_NIGHT);
            sheetStates.add(new SheetExportState(night, nextNo > beforeNo));

            // ИТО
            NoiseSheetCommon.writeSimpleHeader(
                    wb, nonResIto,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_NONRES));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoNonResRoomBlocksFromRow(
                    wb, nonResIto, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_NONRES, true);
            sheetStates.add(new SheetExportState(nonResIto, nextNo > beforeNo));

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoDay,   NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_DAY));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoDay, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_RES_DAY, false);
            int zumStartRow = Math.max(2, resItoDay.getLastRowNum() + 1);
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendZumResRoomBlocksFromRow(
                    wb, resItoDay, building, byKey, zumStartRow, nextNo, thresholds, NoiseTestKind.ZUM_DAY, true);
            sheetStates.add(new SheetExportState(resItoDay, nextNo > beforeNo));

            NoiseSheetCommon.writeSimpleHeader(
                    wb, resItoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.ITO_RES_NIGHT));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseItoSheetWriter.appendItoResRoomBlocksFromRow(
                    wb, resItoNight, building, byKey, 2, nextNo, thresholds, NoiseTestKind.ITO_RES_NIGHT, true);
            sheetStates.add(new SheetExportState(resItoNight, nextNo > beforeNo));

            // Авто
            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoDay, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_DAY));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoDay, building, byKey, 2, nextNo, thresholds, NoiseTestKind.AUTO_DAY);
            sheetStates.add(new SheetExportState(autoDay, nextNo > beforeNo));

            NoiseSheetCommon.writeSimpleHeader(
                    wb, autoNight, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.AUTO_NIGHT));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseAutoSheetWriter.appendAutoRoomBlocksFromRow(
                    wb, autoNight, building, byKey, 2, nextNo, thresholds, NoiseTestKind.AUTO_NIGHT);
            sheetStates.add(new SheetExportState(autoNight, nextNo > beforeNo));

            // Площадка (улица)
            NoiseSheetCommon.writeSimpleHeader(
                    wb, site, NoiseSheetCommon.safeDateLine(dateLines, NoiseTestKind.SITE));
            beforeNo = nextNo;
            nextNo = ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSiteSheetWriter.appendSiteRoomRowsFromRow(
                    wb, site, building, byKey, 2, nextNo, thresholds, NoiseTestKind.SITE);
            sheetStates.add(new SheetExportState(site, nextNo > beforeNo));

            Sheet lastDataSheet = null;
            for (SheetExportState state : sheetStates) {
                if (state.hasData) {
                    lastDataSheet = state.sheet;
                }
            }
            if (lastDataSheet != null) {
                appendFinalNoiseFooter(wb, lastDataSheet);
            }
            removeEmptyNoiseSheets(wb, sheetStates);
            applyNoiseWorkbookFooters(wb);

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

    private static MainFrame findMainFrame(Component parent) {
        if (parent instanceof MainFrame frame) {
            return frame;
        }
        Window window = parent == null ? null : SwingUtilities.getWindowAncestor(parent);
        return (window instanceof MainFrame frame) ? frame : null;
    }

    private static void removeEmptyNoiseSheets(Workbook wb, List<SheetExportState> sheetStates) {
        if (wb == null || sheetStates == null) return;
        for (int i = sheetStates.size() - 1; i >= 0; i--) {
            SheetExportState state = sheetStates.get(i);
            if (state == null || state.hasData || state.sheet == null) continue;
            int index = wb.getSheetIndex(state.sheet);
            if (index >= 0) {
                wb.removeSheetAt(index);
            }
        }
    }

    private static void applyNoiseWorkbookFooters(Workbook wb) {
        if (wb == null || wb.getNumberOfSheets() == 0) return;

        String footerText = buildNoiseFooterText(wb.getSheetAt(0));
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet == null) continue;

            Header header = sheet.getHeader();
            if (header != null) {
                header.setLeft("");
                header.setCenter("");
                header.setRight("");
            }

            Footer footer = sheet.getFooter();
            if (footer != null) {
                footer.setRight(footerText);
            }

            if (i == 0 && sheet instanceof XSSFSheet xssfSheet) {
                CTHeaderFooter headerFooter = xssfSheet.getCTWorksheet().isSetHeaderFooter()
                        ? xssfSheet.getCTWorksheet().getHeaderFooter()
                        : xssfSheet.getCTWorksheet().addNewHeaderFooter();
                headerFooter.setDifferentFirst(true);
                headerFooter.setFirstFooter("");
            }
        }
    }

    private static String buildNoiseFooterText(Sheet titleSheet) {
        String protocolDate = readCellText(titleSheet, 6, 18);
        if (protocolDate.isBlank()) {
            protocolDate = "___ __________ ____ г.";
        }

        String protocolNumber = readCellText(titleSheet, 12, 0);
        String prefix = "Протокол испытаний № ";
        if (protocolNumber.startsWith(prefix)) {
            protocolNumber = protocolNumber.substring(prefix.length()).trim();
        }
        if (protocolNumber.isBlank()) {
            protocolNumber = "_____";
        }

        return "&\"Arial\"&9" +
                "Протокол от " + protocolDate +
                " № " + protocolNumber +
                "\n" +
                "Общее количество страниц &[Страниц] Страница &[Страница]";
    }

    private static String readCellText(Sheet sheet, int rowIndex, int colIndex) {
        if (sheet == null) return "";
        Row row = sheet.getRow(rowIndex);
        if (row == null) return "";
        Cell cell = row.getCell(colIndex);
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell).trim();
    }

    private static void appendFinalNoiseFooter(Workbook wb, Sheet sh) {
        if (wb == null || sh == null) return;

        Font font = wb.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) 10);

        Font boldFont = wb.createFont();
        boldFont.setFontName("Arial");
        boldFont.setFontHeightInPoints((short) 10);
        boldFont.setBold(true);

        CellStyle leftWrap = wb.createCellStyle();
        leftWrap.setFont(font);
        leftWrap.setWrapText(true);
        leftWrap.setAlignment(HorizontalAlignment.LEFT);
        leftWrap.setVerticalAlignment(VerticalAlignment.TOP);

        CellStyle centerWrap = wb.createCellStyle();
        centerWrap.cloneStyleFrom(leftWrap);
        centerWrap.setAlignment(HorizontalAlignment.CENTER);
        centerWrap.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle leftMiddle = wb.createCellStyle();
        leftMiddle.cloneStyleFrom(leftWrap);
        leftMiddle.setVerticalAlignment(VerticalAlignment.CENTER);

        CellStyle signatureLine = wb.createCellStyle();
        signatureLine.cloneStyleFrom(leftMiddle);
        signatureLine.setBorderBottom(BorderStyle.THIN);

        CellStyle centerSmall = wb.createCellStyle();
        centerSmall.cloneStyleFrom(centerWrap);

        CellStyle bottomLine = wb.createCellStyle();
        bottomLine.setFont(font);
        bottomLine.setBorderBottom(BorderStyle.THIN);

        CellStyle protocolEnd = wb.createCellStyle();
        protocolEnd.cloneStyleFrom(centerWrap);
        protocolEnd.setFont(boldFont);

        int row = Math.max(0, sh.getLastRowNum() + 1);
        row = skipExistingRows(sh, row);

        writeMergedText(sh, leftWrap, row, row + 2, 0, 24,
                "Условные обозначения: \n" +
                        "U - значение расширенной неопределенности результата измерения, выраженное в абсолютных значениях. " +
                        "Указанная расширенная неопределенность измерений установлена как стандартная неопределенность измерений, " +
                        "умноженная на коэффициент охвата k (k=2), который соответствует вероятности охвата около 95 %.");
        setRowsHeightPx(sh, row, row + 2, 22f);
        row += 4;

        writeMergedText(sh, leftWrap, row, row + 1, 0, 24,
                "Данные, предоставленные заказчиком, за которые он несет ответственность: нормативные требования в п. 16  " +
                        "Испытательная лаборатория не несет ответственность за данные, предоставленные заказчиком " +
                        "(заявление об ограничении ответственности испытательной лаборатории).");
        setRowsHeightPx(sh, row, row + 1, 22f);
        row += 3;

        writeMergedText(sh, leftMiddle, row, row, 0, 2, "Измерения\u00a0проводил:");
        writeMergedText(sh, signatureLine, row, row, 3, 24,
                "                                             инженер                               Белов Д.А.");
        setRowsHeightPx(sh, row, row, 20f);
        row++;
        writeMergedText(sh, centerSmall, row, row, 0, 21, "(должность, подпись, Ф.И.О.)");
        setRowsHeightPx(sh, row, row, 18f);
        row += 2;

        writeMergedText(sh, leftMiddle, row, row, 0, 2, "Протокол подготовил:");
        writeMergedText(sh, signatureLine, row, row, 3, 24,
                "                                             инженер                               Белов Д.А.");
        setRowsHeightPx(sh, row, row, 20f);
        row++;
        writeMergedText(sh, centerSmall, row, row, 0, 21, "(должность, подпись, Ф.И.О.)");
        setRowsHeightPx(sh, row, row, 18f);
        row += 2;

        writeMergedText(sh, centerWrap, row, row + 4, 0, 24,
                "Испытательная лаборатория несет ответственность за всю информацию, представленную в протоколе испытаний, " +
                        "за исключением случаев, когда информация предоставляется заказчиком.\n" +
                        "Протокол не должен быть воспроизведен не в полном объеме без разрешения испытательной лаборатории ООО «ЦИТ».\n" +
                        "Результаты относятся только к объектам, прошедшим испытания и предоставленным заказчиком.\n" +
                        "Распределение экземпляров протокола испытаний: два протокола – Заказчику, один протокол – испытательной лаборатории ООО «ЦИТ».");
        setRowsHeightPx(sh, row, row + 4, 22f);
        row += 5;

        writeMergedText(sh, bottomLine, row, row, 0, 24, "");
        setRowsHeightPx(sh, row, row, 10f);
        row++;

        writeMergedText(sh, protocolEnd, row, row, 0, 24, "Конец протокола испытаний");
        setRowsHeightPx(sh, row, row, 20f);
    }

    private static int skipExistingRows(Sheet sh, int row) {
        int result = Math.max(0, row);
        while (sh.getRow(result) != null) {
            result++;
        }
        return result;
    }

    private static void setRowsHeightPx(Sheet sh, int firstRow, int lastRow, float heightPx) {
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = NoiseSheetCommon.getOrCreateRow(sh, r);
            row.setHeightInPoints(heightPx * 72f / 96f);
        }
    }

    private static void writeMergedText(Sheet sh, CellStyle style, int firstRow, int lastRow,
                                        int firstCol, int lastCol, String text) {
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        sh.addMergedRegion(region);
        for (int r = firstRow; r <= lastRow; r++) {
            Row row = NoiseSheetCommon.getOrCreateRow(sh, r);
            for (int c = firstCol; c <= lastCol; c++) {
                Cell cell = NoiseSheetCommon.getOrCreateCell(row, c);
                cell.setCellStyle(style);
                if (r == firstRow && c == firstCol) {
                    cell.setCellValue(text);
                }
            }
        }
    }

    private static final class SheetExportState {
        final Sheet sheet;
        final boolean hasData;

        SheetExportState(Sheet sheet, boolean hasData) {
            this.sheet = sheet;
            this.hasData = hasData;
        }
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
