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
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

public final class LightingExcelExporter {

    private LightingExcelExporter() {}

    /**
     * Экспорт Excel «18.2 Естественное освещение».
     *
     * @param building     здание (из текущего проекта)
     * @param sectionIndex индекс секции (0..N-1), либо -1 — все секции
     * @param parent       родительский компонент для диалогов
     */
    public static void export(Building building, int sectionIndex, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = new XSSFWorkbook()) {
            Styles S = new Styles(wb);
            Sheet sh = wb.createSheet("Естественное освещение");

            // ===== ширины столбцов (px → приблиз. ширина Excel) =====
            int[] px = {
                    49, 250, 87, 47, 47, 47, 44, 18, 44, 44, 18, 44, 44, 18, 44, 42, 40, 20, 40
            };
            for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

            // ===== высота строк =====
            // 1 → 16; 2 → 4; 3 → 51; 4 → 160; 5 → 16; 6+ → 16
            ensureRow(sh, 0).setHeightInPoints(16);
            ensureRow(sh, 1).setHeightInPoints(4);
            ensureRow(sh, 2).setHeightInPoints(51);
            ensureRow(sh, 3).setHeightInPoints(160);
            ensureRow(sh, 4).setHeightInPoints(16);

            // ===== A1 =====
            put(sh, 0, 0, "18.2 Естественное освещение:", S.title);

            // ===== Заголовки и объединения =====
            // A3:A4 — № п/п
            mergeWithBorder(sh, "A3:A4");
            put(sh, 2, 0, "№ п/п", S.headCenterBorder);

            // B3:B4 — Наименование места проведения измерений (с переносом)
            mergeWithBorder(sh, "B3:B4");
            put(sh, 2, 1, "Наименование места\nпроведения измерений", S.headCenterBorderWrap);

            // C3:C4 — вертикальный текст, снизу вверх
            mergeWithBorder(sh, "C3:C4");
            put(sh, 2, 2,
                    "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м",
                    S.headVertical);

            // D3:F3 — При верхнем или комбинированном освещении
            mergeWithBorder(sh, "D3:F3");
            put(sh, 2, 3, "При верхнем или комбинированном освещении", S.headCenterBorderWrap);
            put(sh, 3, 3, "Освещенность внутри помещения, лк", S.headVertical);
            put(sh, 3, 4, "Наружная освещенность, лк", S.headVertical);
            put(sh, 3, 5, "КЕО, %", S.headVertical);

            // G3:P3 — При боковом освещении
            mergeWithBorder(sh, "G3:P3");
            put(sh, 2, 6, "При боковом освещении", S.headCenterBorderWrap);
            // G4:I4
            mergeWithBorder(sh, "G4:I4");
            put(sh, 3, 6, "Освещенность внутри помещения ± расширенная неопределенность, лк", S.headVertical);
            // J4:L4
            mergeWithBorder(sh, "J4:L4");
            put(sh, 3, 9, "Наружная освещенность ± расширенная неопределенность, лк", S.headVertical);
            // M4:O4
            mergeWithBorder(sh, "M4:O4");
            put(sh, 3, 12, "КЕО ± расширенная неопределенность, %", S.headVertical);
            // P4
            put(sh, 3, 15, "Допустимое значение КЕО, %", S.headVertical);

            // Q3:S4 — Неравномерность...
            mergeWithBorder(sh, "Q3:S4");
            put(sh, 2, 16, "Неравномерность естественного освещения ± расширенная неопределенность", S.headVertical);

            // ===== Нумерация строки 5 =====
            put(sh, 4, 0, 1, S.headCenterBorder);
            put(sh, 4, 1, 2, S.headCenterBorder);
            put(sh, 4, 2, 3, S.headCenterBorder);
            put(sh, 4, 3, 4, S.headCenterBorder);
            put(sh, 4, 4, 5, S.headCenterBorder);
            put(sh, 4, 5, 6, S.headCenterBorder);
            mergeWithBorder(sh, "G5:I5");     put(sh, 4, 6, 7, S.headCenterBorder);
            mergeWithBorder(sh, "J5:L5");     put(sh, 4, 9, 8, S.headCenterBorder);
            mergeWithBorder(sh, "M5:O5");     put(sh, 4, 12, 9, S.headCenterBorder);
            put(sh, 4, 15, 10, S.headCenterBorder);
            mergeWithBorder(sh, "Q5:S5");     put(sh, 4, 16, 11, S.headCenterBorder);

            // Внешняя рамка для всей шапки (строки 3..5 включительно)
            addRegionBorders(sh, 2, 4, 0, 18);

            // ===== Данные (с 6-й строки, т.е. r=5 0-based) =====
            int row = 5;   // 0-based
            int seq = 1;   // сквозная нумерация колонки A

            List<Entry> entries = collectEntries(building, sectionIndex);
            for (Entry e : entries) {
                // каждая запись занимает блок из 5 строк
                int blockStart = row;
                int blockEnd = row + 4;

                // высоты строк блока
                for (int r = blockStart; r <= blockEnd; r++) ensureRow(sh, r).setHeightInPoints(16);

                // B — объединяем на 5 строк + текст (с периметром)
                mergeWithBorder(sh, "B" + (blockStart + 1) + ":B" + (blockEnd + 1));
                put(sh, blockStart, 1, buildBLabel(building, e), S.centerBorderWrap); // по центру, перенос

                // Q..S — объединяем прямоугольник 3x5 и ставим «-» по центру
                String qRange = "Q" + (blockStart + 1) + ":S" + (blockEnd + 1);
                mergeWithBorder(sh, qRange);
                put(sh, blockStart, 16, "-", S.center);

                boolean isOffice = isOfficeOrPublic(e.space);  // офис/общественное
                String cVal = isOffice ? "Г-0,8" : "Г-0,0";
                double pLimit = isOffice ? 1.00 : 0.50;        // P

                for (int i = 0; i < 5; i++) {
                    int r = row + i;           // 0-based
                    int R = r + 1;             // Excel row (1-based)
                    Row rr = ensureRow(sh, r);

                    // A — сквозная нумерация с гранями
                    Cell A = cell(rr, 0); A.setCellValue(seq++); A.setCellStyle(S.centerBorder);

                    // C — "Г-0,8" / "Г-0,0"
                    Cell C = cell(rr, 2); C.setCellValue(cVal); C.setCellStyle(S.centerBorder);

                    // D,E,F — прочерк "-"
                    for (int c = 3; c <= 5; c++) {
                        Cell x = cell(rr, c); x.setCellValue("-"); x.setCellStyle(S.centerBorder);
                    }

                    // --- Блок G:H:I (внутри, ±, неопределённость) ---
                    Cell G = cell(rr, 6); G.setCellValue(100);                G.setCellStyle(S.num0NoRight);
                    Cell H = cell(rr, 7); H.setCellValue("±");                H.setCellStyle(S.pmTopBottom);
                    Cell I = cell(rr, 8); I.setCellFormula("G" + R + "*0.08*2/POWER(3,0.5)"); I.setCellStyle(S.num0NoLeft);

                    // --- Блок J:K:L (наружная, ±, неопределённость) ---
                    Cell J = cell(rr, 9);  J.setCellValue(6000);              J.setCellStyle(S.num0NoRight);
                    Cell K = cell(rr, 10); K.setCellValue("±");               K.setCellStyle(S.pmTopBottom);
                    Cell L = cell(rr, 11); L.setCellFormula("J" + R + "*0.08*2/POWER(3,0.5)"); L.setCellStyle(S.num0NoLeft);

                    // --- Блок M:N:O (КЕО, ±, неопределённость) ---
                    Cell M = cell(rr, 12); M.setCellFormula("G" + R + "*100/J" + R);           M.setCellStyle(S.num2NoRight);
                    Cell N = cell(rr, 13); N.setCellValue("±");                               N.setCellStyle(S.pmTopBottom);
                    Cell O = cell(rr, 14); O.setCellFormula(
                            "2*M" + R + "*SQRT(POWER(L" + R + "/2/J" + R + ",2)+POWER(I" + R + "/2/G" + R + ",2))"
                    ); O.setCellStyle(S.num2NoLeft);

                    // P — допустимое значение КЕО
                    Cell P = cell(rr, 15); P.setCellValue(pLimit); P.setCellStyle(S.num2);
                }

                row += 5; // следующий блок
            }

            // ===== сохранение =====
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Освещение_естественное.xlsx"));
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

    // ======================================================================
    // Helpers
    // ======================================================================

    /** Внешняя рамка по периметру прямоугольника (включительно). */
    private static void addRegionBorders(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress region = new CellRangeAddress(r1, r2, c1, c2);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);
    }

    /** Объединить диапазон и поставить тонкие границы по периметру. */
    private static void mergeWithBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(r);
        RegionUtil.setBorderTop(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, r, sh);
    }

    private static List<Entry> collectEntries(Building building, int sectionIndex) {
        List<Entry> res = new ArrayList<>();
        for (Floor floor : building.getFloors()) {
            if (floor == null) continue;
            if (sectionIndex >= 0 && floor.getSectionIndex() != sectionIndex) continue;

            for (Space space : floor.getSpaces()) {
                if (space == null) continue;
                for (Room room : space.getRooms()) {
                    if (room != null && room.isSelected()) {
                        res.add(new Entry(floor, space, room));
                    }
                }
            }
        }
        return res;
    }

    private static String buildBLabel(Building building, Entry e) {
        String sect = "";
        try {
            if (building.getSections() != null && !building.getSections().isEmpty()) {
                Section s = building.getSections().get(Math.max(0, e.floor.getSectionIndex()));
                if (s != null && s.getName() != null && !s.getName().isBlank()) sect = s.getName();
            }
        } catch (Exception ignored) {}

        String floorName = (e.floor.getName() != null && !e.floor.getName().isBlank())
                ? e.floor.getName()
                : (e.floor.getNumber() != null ? e.floor.getNumber() : "Этаж");
        String spaceName = spaceDisplayName(e.space);
        String roomName  = e.room.getName() != null ? e.room.getName() : "Комната";

        String label = String.join(", ",
                nonEmpty(sect), nonEmpty(floorName), nonEmpty(spaceName), nonEmpty(roomName));
        if (label.endsWith(", ")) label = label.substring(0, label.length() - 2);
        if (!label.isBlank()) label += ", ";
        label += "нормируемая поверхность";
        return label;
    }

    private static String nonEmpty(String s) { return (s == null) ? "" : s.trim(); }

    private static String spaceDisplayName(Space s) {
        if (s == null) return "Помещение";
        String id = s.getIdentifier();
        if (id != null && !id.isBlank()) return id;
        // если у Space есть getName():
        try {
            java.lang.reflect.Method m = s.getClass().getMethod("getName");
            Object v = m.invoke(s);
            if (v != null) {
                String n = String.valueOf(v);
                if (!n.isBlank()) return n;
            }
        } catch (Exception ignored) {}
        return "Помещение";
    }

    private static boolean isOfficeOrPublic(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("OFFICE") || name.contains("PUBLIC");
    }

    private static Row ensureRow(Sheet sh, int r0) {
        Row r = sh.getRow(r0);
        return (r != null) ? r : sh.createRow(r0);
    }

    private static Cell cell(Row r, int c0) {
        Cell c = r.getCell(c0);
        return (c != null) ? c : r.createCell(c0);
    }

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

    private static void setColWidthPx(Sheet sh, int col, int px) {
        int width = (int) Math.round((px - 5) / 7.0 * 256.0); // примерно как в RadiationExcelExporter
        if (width < 0) width = 0;
        sh.setColumnWidth(col, width);
    }

    // ======================================================================
    // Стили
    // ======================================================================
    private static final class Styles {
        // Шрифты
        final Font arial10;
        final Font arial9;

        // Стили заголовков (строки 3–5) — Arial 9
        final CellStyle headCenterBorder;
        final CellStyle headCenterBorderWrap;
        final CellStyle headVertical;

        // A1 (заголовок листа) и основные стили данных — Arial 10
        final CellStyle title;
        final CellStyle center;
        final CellStyle centerBorder;
        final CellStyle centerBorderWrap;
        final CellStyle textLeftBorderWrap;
        final CellStyle box;
        final CellStyle num0;
        final CellStyle num2;
        final CellStyle pmTopBottom; // для «±»: только верх/низ, без левой/правой грани

        // Стили для подавления боковых граней у соседей «±»
        final CellStyle num0NoRight; // 0 знаков, без правой грани
        final CellStyle num0NoLeft;  // 0 знаков, без левой грани
        final CellStyle num2NoRight; // 2 знака, без правой грани
        final CellStyle num2NoLeft;  // 2 знака, без левой грани

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();

            // === Шрифты ===
            arial10 = wb.createFont();
            arial10.setFontName("Arial");
            arial10.setFontHeightInPoints((short)10);

            arial9 = wb.createFont();
            arial9.setFontName("Arial");
            arial9.setFontHeightInPoints((short)9);

            // === Заголовок A1 — Arial 10 ===
            title = wb.createCellStyle();
            title.setAlignment(HorizontalAlignment.LEFT);
            title.setVerticalAlignment(VerticalAlignment.CENTER);
            title.setFont(arial10);

            // === Шапка (строки 3–5) — Arial 9 ===
            headCenterBorder = wb.createCellStyle();
            headCenterBorder.setAlignment(HorizontalAlignment.CENTER);
            headCenterBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(headCenterBorder);
            headCenterBorder.setFont(arial9);

            headCenterBorderWrap = wb.createCellStyle();
            headCenterBorderWrap.cloneStyleFrom(headCenterBorder);
            headCenterBorderWrap.setWrapText(true);
            headCenterBorderWrap.setFont(arial9);

            headVertical = wb.createCellStyle();
            headVertical.setAlignment(HorizontalAlignment.CENTER);
            headVertical.setVerticalAlignment(VerticalAlignment.CENTER);
            headVertical.setWrapText(true);
            headVertical.setRotation((short) 90); // снизу вверх
            setAllBorders(headVertical);
            headVertical.setFont(arial9);

            // === Данные — Arial 10 ===
            center = wb.createCellStyle();
            center.setAlignment(HorizontalAlignment.CENTER);
            center.setVerticalAlignment(VerticalAlignment.CENTER);
            center.setFont(arial10);

            centerBorder = wb.createCellStyle();
            centerBorder.cloneStyleFrom(center);
            setAllBorders(centerBorder);
            centerBorder.setFont(arial10);

            centerBorderWrap = wb.createCellStyle();
            centerBorderWrap.setAlignment(HorizontalAlignment.CENTER);
            centerBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            centerBorderWrap.setWrapText(true);
            setAllBorders(centerBorderWrap);
            centerBorderWrap.setFont(arial10);

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap);
            textLeftBorderWrap.setFont(arial10);

            box = wb.createCellStyle();
            setAllBorders(box);
            box.setFont(arial10);

            num0 = wb.createCellStyle();
            num0.cloneStyleFrom(centerBorder);
            num0.setDataFormat(fmt.getFormat("0"));
            num0.setFont(arial10);

            num2 = wb.createCellStyle();
            num2.cloneStyleFrom(centerBorder);
            num2.setDataFormat(fmt.getFormat("0.00"));
            num2.setFont(arial10);

            pmTopBottom = wb.createCellStyle();
            pmTopBottom.setAlignment(HorizontalAlignment.CENTER);
            pmTopBottom.setVerticalAlignment(VerticalAlignment.CENTER);
            pmTopBottom.setBorderTop(BorderStyle.THIN);
            pmTopBottom.setBorderBottom(BorderStyle.THIN);
            pmTopBottom.setBorderLeft(BorderStyle.NONE);
            pmTopBottom.setBorderRight(BorderStyle.NONE);
            pmTopBottom.setFont(arial10);

            // num0 без правой/левой грани
            num0NoRight = wb.createCellStyle();
            num0NoRight.cloneStyleFrom(num0);
            num0NoRight.setBorderRight(BorderStyle.NONE);

            num0NoLeft = wb.createCellStyle();
            num0NoLeft.cloneStyleFrom(num0);
            num0NoLeft.setBorderLeft(BorderStyle.NONE);

            // num2 без правой/левой грани
            num2NoRight = wb.createCellStyle();
            num2NoRight.cloneStyleFrom(num2);
            num2NoRight.setBorderRight(BorderStyle.NONE);

            num2NoLeft = wb.createCellStyle();
            num2NoLeft.cloneStyleFrom(num2);
            num2NoLeft.setBorderLeft(BorderStyle.NONE);
        }

        private static void setAllBorders(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }

    private record Entry(Floor floor, Space space, Room room) {}
}
