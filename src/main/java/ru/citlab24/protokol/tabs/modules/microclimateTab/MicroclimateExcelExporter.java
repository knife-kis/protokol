package ru.citlab24.protokol.tabs.modules.microclimateTab;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.Component;
import java.io.File;
import java.io.FileOutputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * Экспорт микроклимата в Excel по заданному шаблону.
 * Логика похожа на LightingExcelExporter, но с другой шапкой.
 */
public final class MicroclimateExcelExporter {

    private MicroclimateExcelExporter() {}

    /** sectionIndex = -1 -> все секции (или одна, если она единственная) */
    public static void export(Building building, int sectionIndex, Component parent) {
        if (building == null) {
            JOptionPane.showMessageDialog(parent, "Сначала загрузите проект (здание).",
                    "Экспорт", JOptionPane.WARNING_MESSAGE);
            return;
        }

        try (Workbook wb = new XSSFWorkbook()) {
            Styles S = new Styles(wb);
            Sheet sh = wb.createSheet("Микроклимат");

            // ===== ширина столбцов (A..V) в пикселях =====
            int[] px = {33,200,28,31,31,34,22,31,45,99,34,14,31,45,33,20,30,45,23,23,23,94};
            for (int c = 0; c < px.length; c++) setColWidthPx(sh, c, px[c]);

            // ===== высота строк (в поинтах) =====
            ensureRow(sh, 0).setHeightInPoints(20);  // 1
            ensureRow(sh, 1).setHeightInPoints(20);  // 2
            ensureRow(sh, 2).setHeightInPoints(48);  // 3
            ensureRow(sh, 3).setHeightInPoints(121); // 4
            ensureRow(sh, 4).setHeightInPoints(20);  // 5

            // ===== заголовки =====
            put(sh, 0, 0, "15. Результаты измерений параметров микроклимата", S.title);
            put(sh, 1, 0, "15.2.  Показатели микроклимата в помещениях:", S.title);

            // ===== шапка =====
            mergeWithBorder(sh, "A3:A4"); put(sh, 2, 0, "№ п/п", S.head8);
            mergeWithBorder(sh, "B3:B4"); put(sh, 2, 1,
                    "Рабочее место, место проведения измерений, цех, участок, наименование профессии или должности",
                    S.head8Wrap);

            mergeWithBorder(sh, "C3:C4"); put(sh, 2, 2, "Дата, проведения  измерений", S.head8Vertical);
            mergeWithBorder(sh, "D3:D4"); put(sh, 2, 3, "Категория работ по интенсивности", S.head8Vertical);
            mergeWithBorder(sh, "E3:E4"); put(sh, 2, 4, "Высота от пола, м", S.head8Vertical);

            mergeWithBorder(sh, "F3:I3"); put(sh, 2, 5, "Температура воздуха, ºС", S.head8);
            mergeWithBorder(sh, "F4:H4"); put(sh, 3, 5, "Измеренная (± расширенная неопределенность)", S.head8Vertical);
            put(sh, 3, 8, "Допустимый уровень", S.head8Vertical);

            put(sh, 2, 9,  "Температура поверхностей, ºС", S.head8);
            put(sh, 3, 9,  "Пол. Измеренная (± расширенная неопределенность)", S.head8Vertical);

            mergeWithBorder(sh, "K3:N3"); put(sh, 2, 10, "Результирующая температура,  ºС", S.head8);
            mergeWithBorder(sh, "K4:M4"); put(sh, 3, 10, "Измеренная (± расширенная неопределенность)", S.head8Vertical);
            put(sh, 3, 13, "Допустимый уровень", S.head8Vertical);

            mergeWithBorder(sh, "O3:R3"); put(sh, 2, 14, "Относительная влажность воздуха, %", S.head8);
            mergeWithBorder(sh, "O4:Q4"); put(sh, 3, 14, "Измеренная (± расширенная неопределенность)", S.head8Vertical);
            put(sh, 3, 17, "Допустимый уровень", S.head8Vertical);

            mergeWithBorder(sh, "S3:V3"); put(sh, 2, 18, "Скорость движения воздуха,  м/с", S.head8);
            mergeWithBorder(sh, "S4:U4"); put(sh, 3, 18, "Измеренная (± расширенная неопределенность)", S.head8Vertical);
            put(sh, 3, 21, "Допустимый уровень", S.head8Vertical);

            // ===== строка 5 (номера граф) =====
            mergeWithBorder(sh, "F5:H5");
            mergeWithBorder(sh, "K5:M5");
            mergeWithBorder(sh, "O5:Q5");
            mergeWithBorder(sh, "S5:U5");
            put(sh, 4, 0,  1, S.centerBorder);
            put(sh, 4, 1,  2, S.centerBorder);
            put(sh, 4, 2,  3, S.centerBorder);
            put(sh, 4, 3,  4, S.centerBorder);
            put(sh, 4, 4,  5, S.centerBorder);
            put(sh, 4, 5,  6, S.centerBorder); // F5-H5 (левый столбец объединённой области)
            put(sh, 4, 8,  7, S.centerBorder);
            put(sh, 4, 9,  8, S.centerBorder);
            put(sh, 4, 10, 9, S.centerBorder); // K5-M5
            put(sh, 4, 13,10, S.centerBorder);
            put(sh, 4, 14,11, S.centerBorder); // O5-Q5
            put(sh, 4, 17,12, S.centerBorder);
            put(sh, 4, 18,13, S.centerBorder); // S5-U5
            put(sh, 4, 21,14, S.centerBorder);

            addRegionBorders(sh, 2, 4, 0, 21); // рамка A3:V5

            // ===== данные =====
            int row = 5; // дальше начинаем вывод с 6-й строки (0-based)
            int seq = 1;
            boolean multipleSections = building.getSections() != null && building.getSections().size() > 1;

            DateTimeFormatter dateFmt = DateTimeFormatter.ofPattern("dd.MM.yyyy");
            String today = LocalDate.now().format(dateFmt);

            // секция -> этаж -> список выбранных комнат
            Map<Integer, Map<Floor, List<Room>>> bySection = new LinkedHashMap<>();
            for (Floor f : building.getFloors()) {
                if (f == null) continue;
                if (sectionIndex >= 0 && f.getSectionIndex() != sectionIndex) continue;

                for (Space s : f.getSpaces()) {
                    if (s == null) continue;

                    List<Room> selected = new ArrayList<>();
                    for (Room r : s.getRooms()) if (r != null && r.isSelected()) selected.add(r);
                    if (selected.isEmpty()) continue;

                    bySection.computeIfAbsent(f.getSectionIndex(), k -> new LinkedHashMap<>())
                            .computeIfAbsent(f, k -> new ArrayList<>())
                            .addAll(selected);
                }
            }

            List<Integer> secOrder = new ArrayList<>(bySection.keySet());
            secOrder.sort(Comparator.naturalOrder());

            for (Integer secIdx : secOrder) {
                Map<Floor, List<Room>> byFloor = bySection.get(secIdx);

                // этажи по возрастанию номера
                List<Floor> floors = new ArrayList<>(byFloor.keySet());
                floors.sort(Comparator.comparingInt(MicroclimateExcelExporter::parseFloorNumSafe));

                for (Floor f : floors) {
                    // заголовок Секция/Этаж (A:V объединено)
                    String title = multipleSections
                            ? joinNotEmpty("Секция " + safeSectionName(building, secIdx), safeFloorName(f))
                            : safeFloorName(f);
                    mergeWithBorder(sh, "A" + (row + 1) + ":V" + (row + 1));
                    put(sh, row, 0, title, S.sectionHeader);
                    row++;

                    // сгруппировать по помещению
                    Map<Space, List<Room>> bySpace = new LinkedHashMap<>();
                    for (Room r : byFloor.get(f)) {
                        Space sp = findRoomSpace(f, r);
                        if (sp == null) continue;
                        bySpace.computeIfAbsent(sp, k -> new ArrayList<>()).add(r);
                    }

                    for (Map.Entry<Space, List<Room>> g : bySpace.entrySet()) {
                        Space space = g.getKey();
                        List<Room> rooms = g.getValue();

                        boolean isResidential = isApartment(space);
                        boolean isOffice      = isOfficeOrPublic(space);

                        for (Room room : rooms) {
                            // 1) центральная точка (3 строки по высотам)
                            row = emitBlock(sh, row, seq++, building, f, space, room,
                                    Position.CENTER, today, isResidential, isOffice, S);

                            // 2) у наружной стены — по числу стен (пока заглушка 0)
                            int walls = externalWallsCount(room); // заменить, когда появятся данные
                            if (isOffice) {
                                if (walls >= 1) {
                                    row = emitBlock(sh, row, seq++, building, f, space, room,
                                            Position.NEAR_WALL, today, isResidential, isOffice, S);
                                }
                            } else {
                                for (int i = 0; i < Math.min(walls, 2); i++) {
                                    row = emitBlock(sh, row, seq++, building, f, space, room,
                                            Position.NEAR_WALL, today, isResidential, isOffice, S);
                                }
                            }
                        }
                    }
                }
            }

            // ===== сохранение =====
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Сохранить Excel");
            chooser.setSelectedFile(new File("Микроклимат.xlsx"));
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
            ex.printStackTrace();
            JOptionPane.showMessageDialog(parent, "Ошибка экспорта: " + ex.getMessage(),
                    "Ошибка", JOptionPane.ERROR_MESSAGE);
        }
    }

    private enum Position { CENTER, NEAR_WALL }

    /** вывод одного блока (3 строки: высоты) */
    private static int emitBlock(Sheet sh, int row, int seq, Building building, Floor f, Space space, Room room,
                                 Position pos, String date, boolean isResidential, boolean isOffice, Styles S) {
        int start = row;

        // A — № п/п (объединяем 3 строки)
        mergeWithBorder(sh, "A" + (start + 1) + ":A" + (start + 3));
        put(sh, start, 0, seq, S.centerBorder);

        // B — место измерения (объединяем 3 строки)
        String roomName  = room.getName() != null ? room.getName() : "";
        String spaceName = spaceDisplayName(space);
        String postfix   = (pos == Position.CENTER) ? "(в центре помещения)" : "(0,5 метра от наружной стены)";
        String where     = joinComma(spaceName, roomName, postfix);
        mergeWithBorder(sh, "B" + (start + 1) + ":B" + (start + 3));
        put(sh, start, 1, where, S.textLeftBorderWrap);

        // C — дата (каждую из 3 строк, вертикально)
        for (int i = 0; i < 3; i++) put(sh, start + i, 2, date, S.vertCenterBorder);

        // D — категория работ (для офисов "3а", для жилья — пусто пока)
        String cat = isOffice ? "3а" : "";
        mergeWithBorder(sh, "D" + (start + 1) + ":D" + (start + 3));
        put(sh, start, 3, cat, S.centerBorder);

        // E — высоты от пола (жильё: 1.5/0.6/0.1; офис: 1.7/0.6/0.1)
        double[] heights = isOffice ? new double[]{1.7, 0.6, 0.1} : new double[]{1.5, 0.6, 0.1};
        for (int i = 0; i < 3; i++) put(sh, start + i, 4, heights[i], S.num1);

        // Остальные графы пока просто в рамке (заполним позже)
        for (int r = 0; r < 3; r++) {
            for (int c = 5; c <= 21; c++) {
                Cell cell = cell(ensureRow(sh, start + r), c);
                if (cell.getCellStyle() == null || cell.getCellStyle().getIndex() == 0) {
                    cell.setCellStyle(S.centerBorder);
                }
            }
            ensureRow(sh, start + r).setHeightInPoints(20); // фиксированная высота строк данных
        }

        return start + 3;
    }

    // ===== утилиты модели/текста =====
    private static String safeSectionName(Building b, int idx) {
        try {
            Section s = b.getSections().get(Math.max(0, idx));
            return s != null && s.getName() != null && !s.getName().isBlank()
                    ? s.getName() : String.valueOf(idx + 1);
        } catch (Exception e) { return String.valueOf(idx + 1); }
    }
    private static String spaceDisplayName(Space s) {
        String id = (s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        return id.isBlank() ? "Помещение" : id;
    }
    private static String safeFloorName(Floor f) {
        String nm = (f.getName() != null && !f.getName().isBlank()) ? f.getName() : f.getNumber();
        return (nm != null) ? nm : "Этаж";
    }
    private static Space findRoomSpace(Floor f, Room r) {
        for (Space s : f.getSpaces()) if (s.getRooms().contains(r)) return s;
        return null;
    }
    private static int parseFloorNumSafe(Floor f) {
        try {
            String s = (f.getNumber() != null) ? f.getNumber() : f.getName();
            if (s == null) return 0;
            java.util.regex.Matcher m = java.util.regex.Pattern.compile("(-?\\d+)").matcher(s);
            return m.find() ? Integer.parseInt(m.group(1)) : 0;
        } catch (Exception e) { return 0; }
    }
    private static boolean isApartment(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("APART"); // APARTMENT/FLAT
    }
    private static boolean isOfficeOrPublic(Space s) {
        if (s == null) return false;
        Space.SpaceType t = s.getType();
        if (t == null) return false;
        String name = t.name().toUpperCase(Locale.ROOT);
        return name.contains("OFFICE") || name.contains("PUBLIC");
    }
    private static int externalWallsCount(Room r) {
        // TODO: заменить когда появится поле в модели (сейчас считаем 0)
        return 0;
    }
    private static String joinComma(String... parts) {
        List<String> xs = new ArrayList<>();
        for (String p : parts) if (p != null && !p.isBlank()) xs.add(p.trim());
        return String.join(", ", xs);
    }
    private static String joinNotEmpty(String... parts) { return joinComma(parts); }

    // ===== табличные утилиты =====
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
    private static void put(Sheet sh, int r0, int c0, double val, CellStyle style) {
        Row r = ensureRow(sh, r0);
        Cell c = cell(r, c0);
        c.setCellValue(val);
        if (style != null) c.setCellStyle(style);
    }
    private static void setColWidthPx(Sheet sh, int col, int px) {
        int width = (int) Math.round((px - 5) / 7.0 * 256.0);
        if (width < 0) width = 0;
        sh.setColumnWidth(col, width);
    }
    private static void addRegionBorders(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress region = new CellRangeAddress(r1, r2, c1, c2);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sh);
    }
    private static void mergeWithBorder(Sheet sh, String addr) {
        CellRangeAddress r = CellRangeAddress.valueOf(addr);
        sh.addMergedRegion(r);
        RegionUtil.setBorderTop(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderBottom(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderLeft(BorderStyle.THIN, r, sh);
        RegionUtil.setBorderRight(BorderStyle.THIN, r, sh);
    }

    // ===== стили =====
    private static final class Styles {
        final Font arial10, arial8;
        final CellStyle title;
        // шапка (8 пт)
        final CellStyle head8, head8Wrap, head8Vertical;
        // данные
        final CellStyle centerBorder, textLeftBorderWrap, num1, vertCenterBorder, sectionHeader;

        Styles(Workbook wb) {
            DataFormat fmt = wb.createDataFormat();

            arial10 = wb.createFont(); arial10.setFontName("Arial"); arial10.setFontHeightInPoints((short)10);
            arial8  = wb.createFont(); arial8.setFontName("Arial");  arial8.setFontHeightInPoints((short)8);

            title = wb.createCellStyle(); title.setFont(arial10);

            head8 = wb.createCellStyle();
            head8.setAlignment(HorizontalAlignment.CENTER);
            head8.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(head8); head8.setFont(arial8);

            head8Wrap = wb.createCellStyle(); head8Wrap.cloneStyleFrom(head8); head8Wrap.setWrapText(true);
            head8Vertical = wb.createCellStyle(); head8Vertical.cloneStyleFrom(head8); head8Vertical.setRotation((short)90);

            centerBorder = wb.createCellStyle();
            centerBorder.setAlignment(HorizontalAlignment.CENTER);
            centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(centerBorder); centerBorder.setFont(arial10);

            textLeftBorderWrap = wb.createCellStyle();
            textLeftBorderWrap.setAlignment(HorizontalAlignment.LEFT);
            textLeftBorderWrap.setVerticalAlignment(VerticalAlignment.CENTER);
            textLeftBorderWrap.setWrapText(true);
            setAllBorders(textLeftBorderWrap); textLeftBorderWrap.setFont(arial10);

            num1 = wb.createCellStyle(); num1.cloneStyleFrom(centerBorder); num1.setDataFormat(fmt.getFormat("0.0"));

            vertCenterBorder = wb.createCellStyle(); vertCenterBorder.cloneStyleFrom(centerBorder); vertCenterBorder.setRotation((short)90);

            sectionHeader = wb.createCellStyle();
            sectionHeader.setAlignment(HorizontalAlignment.LEFT);
            sectionHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            setAllBorders(sectionHeader); sectionHeader.setFont(arial10);
        }

        private static void setAllBorders(CellStyle s) {
            s.setBorderTop(BorderStyle.THIN);
            s.setBorderBottom(BorderStyle.THIN);
            s.setBorderLeft(BorderStyle.THIN);
            s.setBorderRight(BorderStyle.THIN);
        }
    }
}
