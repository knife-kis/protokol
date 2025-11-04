package ru.citlab24.protokol.tabs.modules.lighting;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

/** Экспорт листа «Иск освещение». */
public final class ArtificialLightingExcelExporter {

    private ArtificialLightingExcelExporter() {}

    /* =================== Публичные API =================== */

    public static void export(Building building, int sectionIndex, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        try (Workbook wb = new XSSFWorkbook()) {
            appendToWorkbook(building, sectionIndex, wb);

            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Освещение_искусственное.xlsx"));
            if (chooser.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = chooser.getSelectedFile();
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

    public static void appendToWorkbook(Building building, int sectionIndex, Workbook wb) {
        if (building == null || wb == null) return;
        buildSheet(building, sectionIndex, wb);
    }

    /* =================== Построение листа =================== */

    private static void buildSheet(Building building, int sectionIndex, Workbook wb) {
        Styles S = new Styles(wb);

        Sheet sh = wb.createSheet("Иск освещение");

        // Ориентация/масштаб
        PrintSetup ps = sh.getPrintSetup();
        ps.setLandscape(true);
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        sh.setFitToPage(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);

        // Ширины столбцов (A..P)
        int[] px = {46, 34, 301, 33, 77, 33, 44, 48, 15, 43, 66, 66, 43, 29, 43, 58};
        for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

        // Заголовки
        ensureRow(sh, 0).setHeightInPoints(16);
        ensureRow(sh, 1).setHeightInPoints(16);
        put(sh, 0, 0, "18. Результаты измерений световой среды", S.titleLeft10);
        put(sh, 1, 0, "18.1 Искусственное освещение:",          S.titleLeft10);

        // Высоты строк (px -> pt)
        sh.setDefaultRowHeightInPoints(pxToPt(21)); // последующие
        setRowHeightPx(sh, 2,  4);   // row3
        setRowHeightPx(sh, 3, 45);   // row4
        setRowHeightPx(sh, 4, 58);   // row5
        setRowHeightPx(sh, 5,112);   // row6
        setRowHeightPx(sh, 6, 17);   // row7

        // row3: без «рёбер жёсткости» — только нижняя граница
        for (int c = 0; c <= 15; c++) {
            Cell cc = cell(ensureRow(sh, 2), c);
            cc.setCellStyle(S.bottomOnly);
        }

        // Объединения
        mergeWithBorder(sh, "A4:A6");
        mergeWithBorder(sh, "B4:B6");
        mergeWithBorder(sh, "C4:C6");
        mergeWithBorder(sh, "D4:D6");
        mergeWithBorder(sh, "E4:E6");
        mergeWithBorder(sh, "F4:F6");
        mergeWithBorder(sh, "G4:G6");

        mergeWithBorder(sh, "H4:L4");
        mergeWithBorder(sh, "M4:P4");

        mergeWithBorder(sh, "H5:K5");
        mergeWithBorder(sh, "L5:L6");

        mergeWithBorder(sh, "M5:O5");
        mergeWithBorder(sh, "P5:P6");

        mergeWithBorder(sh, "H7:J7");
        mergeWithBorder(sh, "M7:O7");

        // Подписи шапки (вертикальные там, где «снизу вверх»)
        put(sh, 3, 0, "№ п/п", S.headCenterBorderWrap);

        put(sh, 3, 1, "№ точки измерения", S.headVertical); // B4-B6 снизу вверх
        put(sh, 3, 2, "Наименование места\nпроведения измерений", S.headCenterBorderWrap);
        put(sh, 3, 3, "Разряд, под разряд зрительной работы", S.headVertical); // D4-D6
        put(sh, 3, 4, "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м", S.headVertical); // E4-E6
        put(sh, 3, 5, "Вид, тип светильников", S.headVertical); // F4-F6
        put(sh, 3, 6, "Число не горящих ламп, шт.", S.headVertical); // G4-G6

        put(sh, 3, 7,  "Освещенность, лк", S.headCenterBorderWrap); // H4-L4
        put(sh, 4, 7,  "Измеренная и расширенная неопределенность", S.headCenterBorderWrap); // H5-K5
        put(sh, 5, 10, "общая", S.headVertical); // K6 снизу вверх

        put(sh, 3, 12, "Коэффициент пульсации, %", S.headCenterBorderWrap); // M4-P4
        put(sh, 4, 12, "Измеренная и расширенная неопределенность", S.headCenterBorderWrap); // M5-O5

        put(sh, 4, 11, "Допустимое значение", S.headVertical); // L5-L6 снизу вверх
        put(sh, 4, 15, "Допустимое значение", S.headVertical); // P5-P6 снизу вверх

        // Строка 6 — подписи узких столбцов
        put(sh, 5, 7,  "значение",   S.headVertical); // H6
        put(sh, 5, 8,  "±",          S.headVertical); // I6
        put(sh, 5, 9,  "неопредел.", S.headVertical); // J6
        // K6 уже «общая»
        put(sh, 5, 12, "значение",   S.headVertical); // M6
        put(sh, 5, 13, "±",          S.headVertical); // N6
        put(sh, 5, 14, "неопредел.", S.headVertical); // O6

        // Нумерация строки 7
        int r7 = 6;
        int num = 1;
        for (int c = 0; c <= 11; c++) put(sh, r7, c, num++, S.centerBorder); // A..L
        put(sh, r7, 12, 10, S.centerBorder); // M
        put(sh, r7, 13, 11, S.centerBorder); // N
        put(sh, r7, 14, 12, S.centerBorder); // O

        // ===== ДАННЫЕ (с 8-й строки) =====
        List<Entry> rows = collectSelectedEntries(building, sectionIndex);
        int start = 7; // 0-based row 8
        int seq = 1;

        for (int i = 0; i < rows.size(); i++) {
            Entry e = rows.get(i);
            int r = start + i;
            Row rr = ensureRow(sh, r);

            // A — №
            set(rr, 0, seq++, S.centerBorder);

            // B — прочерк
            set(rr, 1, "-", S.centerBorder);

            // C — "<Этаж>, <Помещение>"
            String cText = floorName(e.floor) + ", " + spaceName(e.space);
            set(rr, 2, cText, S.textLeftBorderWrap);

            // D — прочерк
            set(rr, 3, "-", S.centerBorder);

            // E — "Г-0,8" для офисов/общественных, иначе "Г-0,0"
            String eVal = (isOfficeOrPublic(e.space) ? "Г-0,8" : "Г-0,0");
            set(rr, 4, eVal, S.centerBorder);

            // G — 0
            set(rr, 6, 0, S.centerBorder);

            // H — 100
            set(rr, 7, 100, S.centerBorder);

            // I — ±
            set(rr, 8, "±", S.centerBorder);

            // J — формула
            int R = r + 1; // 1-based
            setFormula(rr, 9, "H" + R + "*0.08*2/POWER(3,0.5)", S.centerBorder);

            // K — прочерк
            set(rr, 10, "-", S.centerBorder);

            // L — прочерк
            set(rr, 11, "-", S.centerBorder);

            // M — пусто
            setBlank(rr, 12, S.centerBorder);

            // N — прочерк
            set(rr, 13, "-", S.centerBorder);

            // O — пусто
            setBlank(rr, 14, S.centerBorder);

            // P — прочерк
            set(rr, 15, "-", S.centerBorder);
        }

        // F8:F(last) — объединяем вниз и пишем вертикально «Общая, светодиодные»
        if (!rows.isEmpty()) {
            int r1 = 8;                     // 1-based первая строка данных
            int r2 = 7 + rows.size();       // 1-based последняя строка данных
            CellRangeAddress reg = new CellRangeAddress(r1 - 1, r2 - 1, 5, 5);
            sh.addMergedRegion(reg);
            RegionUtil.setBorderTop(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderBottom(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderLeft(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderRight(BorderStyle.THIN, reg, sh);
            put(sh, r1 - 1, 5, "Общая, светодиодные", S.headVertical);
        }

        // Разлиновка: шапка (4..7), данные (8..last). Строку 3 НЕ трогаем.
        int lastRow = rows.isEmpty() ? 7 : 7 + rows.size();
        paintGridIfEmpty(sh, 3, 6, 0, 15, S.centerBorder);       // 4..7
        if (lastRow >= 7) paintGridIfEmpty(sh, 7, lastRow, 0, 15, S.centerBorder); // 8..last
    }

    /* =================== Сбор данных =================== */

    private static List<Entry> collectSelectedEntries(Building b, int sectionIndex) {
        List<Entry> res = new ArrayList<>();
        if (b == null) return res;
        for (Floor f : b.getFloors()) {
            if (f == null) continue;
            if (sectionIndex >= 0 && f.getSectionIndex() != sectionIndex) continue;
            for (Space s : f.getSpaces()) {
                if (s == null) continue;
                for (Room r : s.getRooms()) {
                    if (r != null && r.isSelected()) {
                        res.add(new Entry(f, s, r));
                    }
                }
            }
        }
        // По позиции в модели
        res.sort(Comparator.comparingInt((Entry e) -> e.floor.getPosition())
                .thenComparingInt(e -> e.space.getPosition())
                .thenComparingInt(e -> e.room.getPosition()));
        return res;
    }

    private static boolean isOfficeOrPublic(Space s) {
        if (s == null || s.getType() == null) return false;
        String name = s.getType().name().toUpperCase(Locale.ROOT);
        return name.contains("OFFICE") || name.contains("PUBLIC");
    }

    private static String floorName(Floor f) {
        String n = (f != null && f.getName() != null && !f.getName().isBlank())
                ? f.getName() : (f != null && f.getNumber() != null ? f.getNumber() : "Этаж");
        // убрать возможные приставки типа "офисный / жилой" из названия этажа
        return n.replaceFirst("(?iu)^(\\s*)(смешанн(?:ый|ая|ое|ые|ых|ом)?|офисн(?:ый|ая|ое|ые|ых|ом)?|жил(?:ой|ая|ое|ые|ых|ом)?)[\\s,]+", "")
                .replaceAll(" +", " ")
                .trim();
    }

    private static String spaceName(Space s) {
        if (s == null) return "Помещение";
        String id = s.getIdentifier();
        if (id != null && !id.isBlank()) return id;
        try {
            var m = s.getClass().getMethod("getName");
            Object v = m.invoke(s);
            if (v != null) {
                String nm = String.valueOf(v);
                if (!nm.isBlank()) return nm;
            }
        } catch (Exception ignore) {}
        return "Помещение";
    }

    /* =================== Табличные утилиты =================== */

    private static Row ensureRow(Sheet sh, int r0) {
        Row r = sh.getRow(r0);
        return (r != null) ? r : sh.createRow(r0);
    }

    private static Cell cell(Row r, int c0) {
        Cell c = r.getCell(c0);
        return (c != null) ? c : r.createCell(c0);
    }

    // Табличные утилиты (обновлено): перегрузки для строк, целых и вещественных значений
    private static void put(Sheet sh, int r0, int c0, String val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void put(Sheet sh, int r0, int c0, int val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void put(Sheet sh, int r0, int c0, double val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }


    private static void set(Row r, int c0, String val, CellStyle style) {
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void set(Row r, int c0, int val, CellStyle style) {
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }

    private static void setBlank(Row r, int c0, CellStyle style) {
        Cell c = cell(r, c0);
        c.setBlank();
        if (style != null) c.setCellStyle(style);
    }

    private static void setFormula(Row r, int c0, String formula, CellStyle style) {
        Cell c = cell(r, c0);
        c.setCellFormula(formula);
        if (style != null) c.setCellStyle(style);
    }

    private static void setColWidthPx(Sheet sh, int col, int px) {
        int width = (int) Math.round((px - 5) / 7.0 * 256.0);
        if (width < 0) width = 0;
        sh.setColumnWidth(col, width);
    }

    private static void setRowHeightPx(Sheet sh, int r0, int px) {
        ensureRow(sh, r0).setHeightInPoints(pxToPt(px));
    }

    private static float pxToPt(int px) { return px * 0.75f; }

    private static void mergeWithBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(r);
        RegionUtil.setBorderTop(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, r, sh);
    }

    private static void paintGridIfEmpty(Sheet sh, int r1, int r2, int c1, int c2, CellStyle style) {
        for (int rr = r1; rr <= r2; rr++) {
            Row row = ensureRow(sh, rr);
            for (int cc = c1; cc <= c2; cc++) {
                Cell cell = cell(row, cc);
                CellStyle cs = cell.getCellStyle();
                boolean noBorders =
                        cs == null ||
                                (cs.getBorderTop()    == BorderStyle.NONE &&
                                        cs.getBorderBottom() == BorderStyle.NONE &&
                                        cs.getBorderLeft()   == BorderStyle.NONE &&
                                        cs.getBorderRight()  == BorderStyle.NONE);
                if (noBorders) cell.setCellStyle(style);
            }
        }
    }

    /* =================== Стили =================== */

    private static final class Styles {
        final Font arial10, arial9;

        final CellStyle titleLeft10;

        final CellStyle headCenterBorderWrap; // центр + рамка + перенос
        final CellStyle headVertical;         // вертикально (снизу вверх) + рамка
        final CellStyle centerBorder;         // центр + рамка
        final CellStyle textLeftBorderWrap;   // влево + рамка + перенос
        final CellStyle bottomOnly;           // только нижняя граница (для строки 3)

        Styles(Workbook wb) {
            arial10 = wb.createFont(); arial10.setFontName("Arial"); arial10.setFontHeightInPoints((short)10);
            arial9  = wb.createFont();  arial9.setFontName("Arial");  arial9.setFontHeightInPoints((short)9);

            titleLeft10 = wb.createCellStyle();
            titleLeft10.setAlignment(HorizontalAlignment.LEFT);
            titleLeft10.setVerticalAlignment(VerticalAlignment.CENTER);
            titleLeft10.setFont(arial10);

            headCenterBorderWrap = wb.createCellStyle();
            headCenterBorderWrap.setAlignment(HorizontalAlignment.CENTER);
            headCenterBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            headCenterBorderWrap.setWrapText(true);
            setAllBorders(headCenterBorderWrap);
            headCenterBorderWrap.setFont(arial9);

            headVertical = wb.createCellStyle();
            headVertical.setAlignment(HorizontalAlignment.CENTER);
            headVertical.setVerticalAlignment(VerticalAlignment.CENTER);
            headVertical.setWrapText(true);
            headVertical.setRotation((short)90);
            setAllBorders(headVertical);
            headVertical.setFont(arial9);

            centerBorder = wb.createCellStyle();
            centerBorder.setAlignment(HorizontalAlignment.CENTER);
            centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(centerBorder);
            centerBorder.setFont(arial10);

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap);
            textLeftBorderWrap.setFont(arial10);

            bottomOnly = wb.createCellStyle();
            bottomOnly.setAlignment(HorizontalAlignment.LEFT);
            bottomOnly.setVerticalAlignment(VerticalAlignment.CENTER);
            bottomOnly.setBorderTop(BorderStyle.NONE);
            bottomOnly.setBorderLeft(BorderStyle.NONE);
            bottomOnly.setBorderRight(BorderStyle.NONE);
            bottomOnly.setBorderBottom(BorderStyle.THIN);
            bottomOnly.setFont(arial10);
        }

        private static void setAllBorders(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }

    /* =================== Внутренние типы =================== */

    private record Entry(Floor floor, Space space, Room room) {}
}
