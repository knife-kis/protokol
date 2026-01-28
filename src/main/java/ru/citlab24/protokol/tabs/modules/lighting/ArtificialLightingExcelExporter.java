package ru.citlab24.protokol.tabs.modules.lighting;

import org.apache.poi.ss.usermodel.*;
import ru.citlab24.protokol.export.PrintSetupUtils;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.util.*;
import java.util.concurrent.ThreadLocalRandom;

/** Экспорт листа «Иск освещение». */
public final class ArtificialLightingExcelExporter {

    private ArtificialLightingExcelExporter() {}

    /* =================== Публичные API =================== */
    public static void export(Building building,
                              int sectionIndex,
                              Component parent,
                              Map<Integer, Boolean> selectionMap) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        if (selectionMap == null) selectionMap = Collections.emptyMap();

        try (Workbook wb = new XSSFWorkbook()) {
            appendToWorkbook(building, sectionIndex, wb, selectionMap);

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
                    "Экспорт", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void export(Building building, int sectionIndex, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }
        try (Workbook wb = new XSSFWorkbook()) {
            appendToWorkbook(building, sectionIndex, wb, null); // fallback на Room.isSelected()

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

    public static void appendToWorkbook(Building building,
                                        int sectionIndex,
                                        Workbook wb,
                                        Map<Integer, Boolean> selectionMap) {
        if (building == null || wb == null) return;
        buildSheet(building, sectionIndex, wb, selectionMap);
    }


    /* =================== Построение листа =================== */

    private static void buildSheet(Building building,
                                   int sectionIndex,
                                   Workbook wb,
                                   Map<Integer, Boolean> selectionMap) {
        Styles S = new Styles(wb);

        Sheet sh = wb.createSheet("Иск освещение");
        PrintSetup ps = sh.getPrintSetup();
        ps.setLandscape(true);
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        sh.setFitToPage(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0);
        PrintSetupUtils.applyDuplexShortEdge(sh);
        sh.setRepeatingRows(CellRangeAddress.valueOf("7:7"));
        int[] px = {46, 34, 301, 33, 77, 33, 44, 48, 15, 43, 66, 66, 43, 29, 43, 58};
        for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

        ensureRow(sh, 0).setHeightInPoints(16);
        ensureRow(sh, 1).setHeightInPoints(16);
        put(sh, 0, 0, "18. Результаты измерений световой среды", S.titleLeft10);
        put(sh, 1, 0, "18.1 Искусственное освещение:",          S.titleLeft10);

        sh.setDefaultRowHeightInPoints(pxToPt(21));
        setRowHeightPx(sh, 2,  4);
        setRowHeightPx(sh, 3, 45);
        setRowHeightPx(sh, 4, 58);
        setRowHeightPx(sh, 5,112);
        setRowHeightPx(sh, 6, 17);

        // строка 3 — только нижняя линия
        for (int c = 0; c <= 15; c++) cell(ensureRow(sh, 2), c).setCellStyle(S.bottomOnly);

        // объединения
        mergeWithBorder(sh, "A4:A6");
        mergeWithBorder(sh, "B4:B6");
        mergeWithBorder(sh, "C4:C6");
        mergeWithBorder(sh, "D4:D6");
        mergeWithBorder(sh, "E4:E6");
        mergeWithBorder(sh, "F4:F6");
        mergeWithBorder(sh, "G4:G6");

        mergeWithBorder(sh, "H4:L4");
        mergeWithBorder(sh, "M4:P4");

        mergeWithBorder(sh, "H5:K5");   // «Измеренная ± расширенная неопределенность»
        mergeWithBorder(sh, "L5:L6");   // «Нормируемая»
        mergeWithBorder(sh, "M5:O6");   // «Измеренный ± расширенная неопределенность»
        mergeWithBorder(sh, "P5:P6");   // «Нормируемая»

        mergeWithBorder(sh, "H6:J6");   // «Общая»
        RegionUtil.setBorderTop(BorderStyle.THIN,   new CellRangeAddress(5,5,10,10), sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN,new CellRangeAddress(5,5,10,10), sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN,  new CellRangeAddress(5,5,10,10), sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, new CellRangeAddress(5,5,10,10), sh);

        mergeWithBorder(sh, "H7:J7");
        mergeWithBorder(sh, "M7:O7");

        // подписи шапки
        put(sh, 3, 0, "№ п/п", S.headCenterBorderWrap);
        put(sh, 3, 1, "№ точки измерения", S.headVertical);
        put(sh, 3, 2, "Наименование места\nпроведения измерений", S.headCenterBorderWrap);
        put(sh, 3, 3, "Разряд, под разряд зрительной работы", S.headVertical);
        put(sh, 3, 4, "Рабочая поверхность, плоскость измерения (горизонтальная - Г, вертикальная - В) - высота от пола (земли), м", S.headVertical);
        put(sh, 3, 5, "Вид, тип светильников", S.headVertical);
        put(sh, 3, 6, "Число не горящих ламп, шт.", S.headVertical);

        put(sh, 3, 7,  "Освещенность, лк", S.headCenterBorderWrap);
        put(sh, 4, 7,  "Измеренная ± расширенная неопределенность", S.headCenterBorderWrap);
        put(sh, 5, 7,  "Общая", S.headCenterBorderWrap);
        put(sh, 5, 10, "Комбинированная", S.headVertical);
        put(sh, 4, 11, "Нормируемая", S.headVertical);

        put(sh, 3, 12, "Коэффициент пульсации, %", S.headCenterBorderWrap);
        put(sh, 4, 12, "Измеренный ± расширенная неопределенность", S.headCenterBorderWrap);
        put(sh, 4, 15, "Нормируемая", S.headVertical);

        // строка 7 — 1..12
        int r7 = 6;
        int k = 1;
        for (int c = 0; c <= 6; c++) put(sh, r7, c, k++, S.centerBorder); // A..G = 1..7
        put(sh, r7, 7, 8,  S.centerBorder);  // H..J
        put(sh, r7, 10, 9, S.centerBorder);  // K
        put(sh, r7, 11, 10,S.centerBorder);  // L
        put(sh, r7, 12, 11,S.centerBorder);  // M..O
        put(sh, r7, 15, 12,S.centerBorder);  // P

        // ===== ДАННЫЕ (с 8-й строки) =====
        List<Entry> rows = collectSelectedEntries(building, sectionIndex, selectionMap);
        System.out.println("[ArtificialLighting] rooms to export = " + rows.size());
        int start = 7; // 0-based (row 8)
        int seq = 1;

        for (int i = 0; i < rows.size(); i++) {
            Entry e = rows.get(i);
            int r = start + i;
            Row rr = ensureRow(sh, r);

            set(rr, 0, seq++, S.centerBorder);       // A
            set(rr, 1, "-",  S.centerBorder);        // B

            String roomName = (e.room.getName() != null && !e.room.getName().isBlank())
                    ? e.room.getName().trim() : "Комната";
            set(rr, 2, floorName(e.floor) + ", " + roomName, S.textLeftBorderWrap); // C

            set(rr, 3, "-",  S.centerBorder);        // D

            boolean isOffice = isOfficeSpace(e.space);
            String eVal = isOffice ? "Г-0,8" : "Г-0,0";
            set(rr, 4, eVal, S.centerBorder);        // E

            setBlank(rr, 5, S.centerBorder);         // F
            set(rr, 6, 0, S.centerBorder);           // G

            // ---------- H/I/J: значение + «±» + погрешность ----------
            Integer lNormByName = normativeL(roomName); // 20/30 или null
            double hVal;
            if (isOffice) {
                // офисы: 500..610
                hVal = randomWithDecimals(500, 610, 0);
            } else if (lNormByName != null && lNormByName == 20) {
                hVal = randomWithDecimals(49, 105, 1);   // 49.0 .. 105.0 (шаг 0.1)
            } else if (lNormByName != null && lNormByName == 30) {
                hVal = randomWithDecimals(110, 175, 1);  // 110.0 .. 175.0 (шаг 0.1)
            } else {
                hVal = 100;
            }

            // Формат H/J: если H >=100 → 0 знаков; если 10..99.9 → 1 знак.
            boolean use0dec = (hVal >= 100);
            CellStyle hStyle = use0dec ? S.num0NoRightRose : S.num1NoRightRose;
            CellStyle jStyle = use0dec ? S.num0NoLeft     : S.num1NoLeft;

            // H (с розовой подложкой, без правой грани)
            set(rr, 7, hVal, hStyle);

            // I: «±» без боковых граней
            set(rr, 8, "±", S.pmTopBottom);

            // J: формула от H, стиль под точность, без левой грани
            int R1 = r + 1;
            setFormula(rr, 9, "H" + R1 + "*0.08*2/POWER(3,0.5)", jStyle);

            // ---------- K / L ----------
            set(rr, 10, "-", S.centerBorder);  // K

// L: НОРМАТИВНАЯ освещённость (без авто-распаковки)
            if (isOffice) {
                set(rr, 11, 300, S.centerBorder);                           // офисам всегда 300
            } else if (lNormByName != null) {
                set(rr, 11, lNormByName.intValue(), S.centerBorder);        // 20 или 30
            } else {
                set(rr, 11, "-", S.centerBorder);                           // не распознали — «-»
            }

            // ---------- M / N / O (коэфф. пульсации) ----------
            if (isOffice) {
                double p = ThreadLocalRandom.current().nextDouble();
                double mVal = (p < 0.90) ? 1.1 : (p < 0.98 ? 1.2 : 1.3);
                Cell mCell = cell(rr, 12);
                mCell.setCellValue(mVal);
                mCell.setCellStyle(S.num1NoRight);   // 0.0, без правой
            } else {
                setBlank(rr, 12, S.centerNoRight);
            }

            set(rr, 13, "-",  S.pmTopBottom);  // N (± без боковых)
            setBlank(rr, 14, S.centerNoLeft);  // O

            set(rr, 15, "-",  S.centerBorder); // P
        }

        // F8:F(last) — объединяем и подписываем вертикально
        if (!rows.isEmpty()) {
            int r1 = 8;                   // 1-based
            int r2 = 7 + rows.size();     // 1-based
            CellRangeAddress reg = new CellRangeAddress(r1 - 1, r2 - 1, 5, 5);
            sh.addMergedRegion(reg);
            RegionUtil.setBorderTop(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderBottom(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderLeft(BorderStyle.THIN, reg, sh);
            RegionUtil.setBorderRight(BorderStyle.THIN, reg, sh);
            put(sh, r1 - 1, 5, "Общая, светодиодные", S.headVertical);
        }

        // Разлиновка: БЕЗ лишней последней строки
        int lastRow = rows.isEmpty() ? 6 : (start + rows.size() - 1); // <-- фикс
        paintGridIfEmpty(sh, 3, 6, 0, 15, S.centerBorder);           // 4..7
        if (lastRow >= 7) paintGridIfEmpty(sh, 7, lastRow, 0, 15, S.centerBorder); // 8..last
    }

    /* =================== Сбор данных =================== */

    private static List<Entry> collectSelectedEntries(Building b, int sectionIndex, Map<Integer, Boolean> selectionMap) {
        List<Entry> res = new ArrayList<>();
        if (b == null) return res;

        boolean useMap = (selectionMap != null && !selectionMap.isEmpty());

        for (Floor f : b.getFloors()) {
            if (f == null) continue;
            if (sectionIndex >= 0 && f.getSectionIndex() != sectionIndex) continue;

            for (Space s : f.getSpaces()) {
                if (s == null) continue;
                // Экспортируем только офисные/общественные
                if (!(isOfficeSpace(s) || isPublicSpace(s))) continue;

                for (Room r : s.getRooms()) {
                    if (r == null) continue;
                    boolean selected = isRoomSelectedForArtificial(r, useMap ? selectionMap : null);
                    if (selected) res.add(new Entry(f, s, r));
                }
            }
        }

        res.sort(Comparator.comparingInt((Entry e) -> e.floor.getPosition())
                .thenComparingInt(e -> e.space.getPosition())
                .thenComparingInt(e -> e.room.getPosition()));
        return res;
    }

    private static Integer normativeL(String roomName) {
        if (roomName == null) return null;
        String s = roomName.toLowerCase(Locale.ROOT);

        // 20 лк
        String[] p20 = {
                "лестничная клетка", "лестница", "лифтовой холл",
                "тамбур", "коридор", "техническ", "колясочн"
        };
        for (String p : p20) if (s.contains(p)) return 20;

        // 30 лк
        String[] p30 = {
                "электрощитов", "электрощитовая",    // Электрощитовая
                "венткамера",                        // Венткамера
                "итп",                               // ИТП
                "пнс",                               // ПНС
                "водомерн",                          // водомерный узел
                "узел учета", "узел учёта",          // узел учета/учёта
                "тепловой пункт",                    // тепловой пункт
                "машинное помещение",                // машинное помещение
                "узел ввода",                        // узел ввода
                "насосная",                          // насосная
                "вестибюл"                           // вестибюль
        };
        for (String p : p30) if (s.contains(p)) return 30;

        return null; // не распознали
    }


    private static boolean isOfficeSpace(Space s) {
        if (s == null || s.getType() == null) return false;
        return s.getType().name().toUpperCase(Locale.ROOT).contains("OFFICE");
    }
    private static boolean isPublicSpace(Space s) {
        if (s == null || s.getType() == null) return false;
        return s.getType().name().toUpperCase(Locale.ROOT).contains("PUBLIC");
    }

    private static String floorName(Floor f) {
        String n = (f != null && f.getName() != null && !f.getName().isBlank())
                ? f.getName()
                : (f != null && f.getNumber() != null ? f.getNumber() : "Этаж");
        // дополнительно убираем префикс «общественный …»
        return n.replaceFirst("(?iu)^(\\s*)(смешанн(?:ый|ая|ое|ые|ых|ом)?|офисн(?:ый|ая|ое|ые|ых|ом)?|жил(?:ой|ая|ое|ые|ых|ом)?|общественн(?:ый|ая|ое|ые|ых|ом)?)[\\s,]+", "")
                .replaceAll(" +", " ")
                .trim();
    }
    /** Определяем, отмечена ли комната на вкладке «Искусственное освещение».
     * Приоритет: карта selectionMap -> спец-флаг в Room (isArtificialLightingSelected / getArtificialLightingSelected / isArtificialSelected)
     * -> общий Room.isSelected() как последний фолбэк.
     */
    private static boolean isRoomSelectedForArtificial(Room r, Map<Integer, Boolean> selectionMap) {
        if (r == null) return false;

        // 1) Карта из вкладки (если передана и id валиден)
        if (selectionMap != null && !selectionMap.isEmpty()) {
            try {
                Boolean v = selectionMap.get(r.getId());
                if (v != null) return v;
            } catch (Throwable ignore) {}
        }

        // 2) Только специальные флаги «искусственного» в модели Room
        try {
            Object v = r.getClass().getMethod("isArtificialSelected").invoke(r);
            if (v instanceof Boolean) return (Boolean) v;
        } catch (Throwable ignore) {}
        try {
            Object v = r.getClass().getMethod("isArtificialLightingSelected").invoke(r);
            if (v instanceof Boolean) return (Boolean) v;
        } catch (Throwable ignore) {}
        try {
            Object v = r.getClass().getMethod("getArtificialLightingSelected").invoke(r);
            if (v instanceof Boolean) return (Boolean) v;
        } catch (Throwable ignore) {}

        // ВАЖНО: больше НЕ падаем на Room.isSelected() (КЕО) — это другой модуль
        return false;
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

    private static void set(Row r, int c0, double val, CellStyle style) {
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

    private static double randomWithDecimals(int minInclusive, int maxInclusive, int decimals) {
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        if (decimals <= 0) {
            return rnd.nextInt(minInclusive, maxInclusive + 1);
        }
        int scale = (int) Math.pow(10, decimals);
        int minScaled = minInclusive * scale;
        int maxScaled = maxInclusive * scale;
        int valueScaled = rnd.nextInt(minScaled, maxScaled + 1);
        return valueScaled / (double) scale;
    }

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
        final CellStyle headCenterBorderWrap;
        final CellStyle headVertical;
        final CellStyle centerBorder;
        final CellStyle textLeftBorderWrap;
        final CellStyle bottomOnly;
        final CellStyle pmTopBottom;       // только верх+низ
        final CellStyle centerNoRight;     // центр, нет правой
        final CellStyle centerNoLeft;      // центр, нет левой
        final CellStyle centerNoRightRose; // как centerNoRight, но нежно-красная заливка (H)
        final CellStyle num1NoRight;       // 0.0, без правой (M)
        final CellStyle num0NoRightRose;   // 0, без правой + розовая (H  ≥100)
        final CellStyle num1NoRightRose;   // 0.0, без правой + розовая (10..99.9)
        final CellStyle num0NoLeft;        // 0, без левой   (J для H ≥100)
        final CellStyle num1NoLeft;        // 0.0, без левой (J для H 10..99.9)

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();

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
            headCenterBorderWrap.setFont(arial10);

            headVertical = wb.createCellStyle();
            headVertical.setAlignment(HorizontalAlignment.CENTER);
            headVertical.setVerticalAlignment(VerticalAlignment.CENTER);
            headVertical.setWrapText(true);
            headVertical.setRotation((short)90);
            setAllBorders(headVertical);
            headVertical.setFont(arial10);

            centerBorder = wb.createCellStyle();
            centerBorder.setAlignment(HorizontalAlignment.CENTER);
            centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(centerBorder);
            centerBorder.setFont(arial9);

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap);
            textLeftBorderWrap.setFont(arial9);

            bottomOnly = wb.createCellStyle();
            bottomOnly.setAlignment(HorizontalAlignment.LEFT);
            bottomOnly.setVerticalAlignment(VerticalAlignment.CENTER);
            bottomOnly.setBorderTop(BorderStyle.NONE);
            bottomOnly.setBorderLeft(BorderStyle.NONE);
            bottomOnly.setBorderRight(BorderStyle.NONE);
            bottomOnly.setBorderBottom(BorderStyle.THIN);
            bottomOnly.setFont(arial9);

            pmTopBottom = wb.createCellStyle();
            pmTopBottom.setAlignment(HorizontalAlignment.CENTER);
            pmTopBottom.setVerticalAlignment(VerticalAlignment.CENTER);
            pmTopBottom.setBorderTop(BorderStyle.THIN);
            pmTopBottom.setBorderBottom(BorderStyle.THIN);
            pmTopBottom.setBorderLeft(BorderStyle.NONE);
            pmTopBottom.setBorderRight(BorderStyle.NONE);
            pmTopBottom.setFont(arial9);

            centerNoRight = wb.createCellStyle();
            centerNoRight.setAlignment(HorizontalAlignment.CENTER);
            centerNoRight.setVerticalAlignment(VerticalAlignment.CENTER);
            centerNoRight.setBorderTop(BorderStyle.THIN);
            centerNoRight.setBorderBottom(BorderStyle.THIN);
            centerNoRight.setBorderLeft(BorderStyle.THIN);
            centerNoRight.setBorderRight(BorderStyle.NONE);
            centerNoRight.setFont(arial9);

            centerNoLeft = wb.createCellStyle();
            centerNoLeft.setAlignment(HorizontalAlignment.CENTER);
            centerNoLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            centerNoLeft.setBorderTop(BorderStyle.THIN);
            centerNoLeft.setBorderBottom(BorderStyle.THIN);
            centerNoLeft.setBorderRight(BorderStyle.THIN);
            centerNoLeft.setBorderLeft(BorderStyle.NONE);
            centerNoLeft.setFont(arial9);

// H: базовая розовая подложка
            centerNoRightRose = wb.createCellStyle();
            centerNoRightRose.cloneStyleFrom(centerNoRight);
            centerNoRightRose.setFillForegroundColor(IndexedColors.ROSE.getIndex());
            centerNoRightRose.setFillPattern(FillPatternType.SOLID_FOREGROUND);

// H: точные варианты
            num0NoRightRose = wb.createCellStyle();
            num0NoRightRose.cloneStyleFrom(centerNoRightRose);
            num0NoRightRose.setDataFormat(fmt.getFormat("0"));

            num1NoRightRose = wb.createCellStyle();
            num1NoRightRose.cloneStyleFrom(centerNoRightRose);
            num1NoRightRose.setDataFormat(fmt.getFormat("0.0"));

// J: точные варианты «без левой»
            num0NoLeft = wb.createCellStyle();
            num0NoLeft.cloneStyleFrom(centerNoLeft);
            num0NoLeft.setDataFormat(fmt.getFormat("0"));

            num1NoLeft = wb.createCellStyle();
            num1NoLeft.cloneStyleFrom(centerNoLeft);
            num1NoLeft.setDataFormat(fmt.getFormat("0.0"));

// M: число с 1 знаком, без правой грани
            num1NoRight = wb.createCellStyle();
            num1NoRight.cloneStyleFrom(centerNoRight);
            num1NoRight.setDataFormat(fmt.getFormat("0.0"));

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
