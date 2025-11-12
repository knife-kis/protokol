package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import static ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon.*;

public class NoiseLiftSheetWriter {

    private NoiseLiftSheetWriter() {}

    /** Полная шапка лифтового листа (дневной). */
    public static void writeLiftHeader(Workbook wb, Sheet sh, String dateLine) {
        Font f8  = wb.createFont();  f8.setFontName("Arial");  f8.setFontHeightInPoints((short) 8);
        Font f10 = wb.createFont();  f10.setFontName("Arial"); f10.setFontHeightInPoints((short)10);
        Font f7  = wb.createFont();  f7.setFontName("Arial");  f7.setFontHeightInPoints((short) 7);

        CellStyle center = wb.createCellStyle();
        center.setAlignment(HorizontalAlignment.CENTER);
        center.setVerticalAlignment(VerticalAlignment.CENTER);
        center.setWrapText(false);
        center.setFont(f8);

        CellStyle centerWrap = wb.createCellStyle();
        centerWrap.cloneStyleFrom(center);
        centerWrap.setWrapText(true);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.cloneStyleFrom(center);
        setThinBorder(centerBorder);

        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerWrap);
        setThinBorder(centerWrapBorder);

        CellStyle leftNoWrap = wb.createCellStyle();
        leftNoWrap.setAlignment(HorizontalAlignment.LEFT);
        leftNoWrap.setVerticalAlignment(VerticalAlignment.CENTER);
        leftNoWrap.setWrapText(false);
        leftNoWrap.setFont(f8);

        CellStyle verticalBorder = wb.createCellStyle();
        verticalBorder.cloneStyleFrom(centerBorder);
        verticalBorder.setRotation((short)90);

        CellStyle leftNoWrapF10 = wb.createCellStyle();
        leftNoWrapF10.cloneStyleFrom(leftNoWrap);
        leftNoWrapF10.setFont(f10);

        CellStyle centerBorderF10 = wb.createCellStyle();
        centerBorderF10.cloneStyleFrom(centerBorder);
        centerBorderF10.setFont(f10);

        CellStyle centerBorderF10_noWrap = wb.createCellStyle();
        centerBorderF10_noWrap.cloneStyleFrom(centerBorder);
        centerBorderF10_noWrap.setWrapText(false);
        centerBorderF10_noWrap.setFont(f10);

        CellStyle centerWrapBorderF7 = wb.createCellStyle();
        centerWrapBorderF7.cloneStyleFrom(centerWrapBorder);
        centerWrapBorderF7.setFont(f7);

        CellStyle centerWrapBorderF10 = wb.createCellStyle();
        centerWrapBorderF10.cloneStyleFrom(centerWrapBorder);
        centerWrapBorderF10.setFont(f10);

        CellStyle verticalBorderF10 = wb.createCellStyle();
        verticalBorderF10.cloneStyleFrom(verticalBorder);
        verticalBorderF10.setFont(f10);

        CellStyle verticalWrapBorderF10 = wb.createCellStyle();
        verticalWrapBorderF10.cloneStyleFrom(verticalBorderF10);
        verticalWrapBorderF10.setWrapText(true);

        sh.setDefaultRowHeightInPoints(cmToPt(0.53));
        setRowHeightCm(sh, 0, 0.53);
        setRowHeightCm(sh, 1, 0.53);
        setRowHeightCm(sh, 2, 0.53);
        setRowHeightCm(sh, 3, 0.74);
        setRowHeightCm(sh, 4, 2.94);
        setRowHeightCm(sh, 5, 0.53);

        Row r1 = getOrCreateRow(sh, 0);
        Cell a1 = getOrCreateCell(r1, 0);
        a1.setCellValue("16. Результаты измерений виброакустических факторов");
        a1.setCellStyle(leftNoWrapF10);

        Row r2 = getOrCreateRow(sh, 1);
        Cell a2 = getOrCreateCell(r2, 0);
        a2.setCellValue("16.2. Шум:");
        a2.setCellStyle(leftNoWrapF10);

        java.util.List<CellRangeAddress> merges = new java.util.ArrayList<>();

        merges.add(merge(sh, 2, 4, 0, 0));
        setCenter(sh, 2, 0, "№ п/п", verticalBorderF10);

        merges.add(merge(sh, 2, 4, 1, 1));
        setCenter(sh, 2, 1, "№ точки измерения", verticalBorderF10);

        merges.add(merge(sh, 2, 4, 2, 2));
        setCenter(sh, 2, 2, "Место измерений", centerBorderF10_noWrap);

        merges.add(merge(sh, 2, 4, 3, 3));
        setCenter(sh, 2, 3, "Источник шума\n(тип, вид, марка, условия замера)", centerWrapBorderF10);

        merges.add(merge(sh, 2, 2, 4, 8));
        setCenter(sh, 2, 4, "Характер шума", centerBorderF10_noWrap);

        merges.add(merge(sh, 3, 3, 4, 5));
        setCenter(sh, 3, 4, "по спектру", centerWrapBorderF7);

        merges.add(merge(sh, 3, 3, 6, 8));
        setCenter(sh, 3, 6, "по временным характеристикам", centerWrapBorderF7);

        Row r5 = getOrCreateRow(sh, 4);
        setText(r5, 4, "широкополосный", verticalBorderF10);
        setText(r5, 5, "тональный",       verticalBorderF10);
        setText(r5, 6, "постоянный",      verticalBorderF10);
        setText(r5, 7, "непостоянный",    verticalBorderF10);
        setText(r5, 8, "импульсный",      verticalBorderF10);

        merges.add(merge(sh, 2, 3, 9, 17));
        setCenter(sh, 2, 9,
                "Уровни звукового давления (дБ) ± U (дБ) в октавных полосах частот со среднегеометрическими частотами (Гц)",
                centerWrapBorder);

        String[] freqs = {"31,5","63","125","250","500","1000","2000","4000","8000"};
        for (int i = 0; i < freqs.length; i++) setText(r5, 9 + i, freqs[i], centerBorderF10);

        merges.add(merge(sh, 2, 4, 18, 18));
        setCenter(sh, 2, 18, "Уровни звука (дБА) ±U (дБ)", verticalWrapBorderF10);

        merges.add(merge(sh, 2, 4, 19, 21));
        setCenter(sh, 2, 19, "Эквивалентные уровни звука,  (дБА) ±U (дБ)", verticalWrapBorderF10);

        merges.add(merge(sh, 2, 4, 22, 24));
        setCenter(sh, 2, 22, "Максимальные уровни звука  (дБА)", verticalWrapBorderF10);

        Row r6 = getOrCreateRow(sh, 5);
        for (int c = 0; c <= 18; c++) setText(r6, c, String.valueOf(c + 1), (c <= 3) ? centerBorderF10 : centerBorder);
        merges.add(merge(sh, 5, 5, 19, 21)); setCenter(sh, 5, 19, "20", centerBorder);
        merges.add(merge(sh, 5, 5, 22, 24)); setCenter(sh, 5, 22, "21", centerBorder);

        for (CellRangeAddress rgn : merges) {
            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rgn, sh);
        }

        if (dateLine != null && !dateLine.isBlank()) {
            setRowHeightCm(sh, 6, 0.53);
            CellRangeAddress rng = merge(sh, 6, 6, 0, 24);
            Row r7 = getOrCreateRow(sh, 6);
            Cell a7 = getOrCreateCell(r7, 0);
            a7.setCellValue(dateLine);

            CellStyle centerBorder8 = wb.createCellStyle();
            centerBorder8.cloneStyleFrom(centerBorder);
            centerBorder8.setFont(f8);
            centerBorder8.setAlignment(HorizontalAlignment.CENTER);
            centerBorder8.setVerticalAlignment(VerticalAlignment.CENTER);
            a7.setCellStyle(centerBorder8);

            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rng, sh);
        }
    }

    /** Добавляет блоки «т1/т2/т3» начиная со строки 7 (дневной) / с указанной строки (ночной). */
    public static void appendLiftRoomBlocks(Workbook wb, Sheet sh,
                                            Building building,
                                            java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                            boolean skipOfficeAtNight) {
        appendLiftRoomBlocksFromRow(wb, sh, building, byKey, 7, skipOfficeAtNight);
    }

    public static void appendLiftRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                   Building building,
                                                   java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                   int startRow,
                                                   boolean skipOfficeAtNight) {
        if (building == null) return;

        org.apache.poi.ss.usermodel.Font f8 = wb.createFont();
        f8.setFontName("Arial");
        f8.setFontHeightInPoints((short)8);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.setAlignment(HorizontalAlignment.CENTER);
        centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
        centerBorder.setWrapText(false);
        centerBorder.setFont(f8);
        setThinBorder(centerBorder);

        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerBorder);
        centerWrapBorder.setWrapText(true); // для D: перенос строк

        CellStyle leftNoWrapBorder = wb.createCellStyle();
        leftNoWrapBorder.cloneStyleFrom(centerBorder);
        leftNoWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftNoWrapBorder.setWrapText(false);

        // НОВОЕ: стиль с переносом для C (место измерений)
        CellStyle leftWrapBorder = wb.createCellStyle();
        leftWrapBorder.cloneStyleFrom(centerBorder);
        leftWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftWrapBorder.setWrapText(true);

        final String[] PLUS_MINUS = { "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-" };

        int row = Math.max(0, startRow);
        int no  = 1;

        java.util.List<Section> sections = building.getSections();
        boolean multiSections = sections != null && sections.size() > 1;

        java.util.List<Floor> floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor fl : floors) {
            String floorNum = (fl.getNumber() == null) ? "" : fl.getNumber().trim();

            java.util.List<Space> spaces = new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space sp : spaces) {
                if (skipOfficeAtNight && sp != null && sp.getType() == Space.SpaceType.OFFICE) continue;

                String spaceId = (sp.getIdentifier() == null) ? "" : sp.getIdentifier().trim();

                java.util.List<Room> rooms = new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room rm : rooms) {
                    String roomName = (rm.getName() == null) ? "" : rm.getName().trim();
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null || !nv.lift) continue;

                    Section sec = null;
                    if (multiSections && fl.getSectionIndex() >= 0 && fl.getSectionIndex() < sections.size()) {
                        sec = sections.get(fl.getSectionIndex());
                    }

                    String place = formatPlace(building, sec, fl, sp, rm);

                    for (int t = 1; t <= 3; t++) {
                        int r1 = row, r2 = row + 1, r3 = row + 2;

                        // НОВОЕ: первая строка — авто-высота (Excel сам подберёт по переносам)
                        Row R1 = getOrCreateRow(sh, r1);
                        R1.setHeight((short)-1); // авто

                        // как и было: фиксированные высоты для 2-й и 3-й строк
                        setRowHeightCm(sh, r2, 0.53);
                        setRowHeightCm(sh, r3, 0.53);

                        // базовая сетка A..Y
                        for (int rr = r1; rr <= r3; rr++) {
                            Row cur = getOrCreateRow(sh, rr);
                            for (int c = 0; c <= 24; c++) {
                                Cell cell = getOrCreateCell(cur, c);
                                cell.setCellStyle(centerBorder);
                            }
                        }

                        // A: № п/п (merge 3 строки)
                        CellRangeAddress aMerge = merge(sh, r1, r3, 0, 0);
                        setCenter(sh, r1, 0, String.valueOf(no++), centerBorder);

                        // B: т1/т2/т3 (merge 3 строки)
                        CellRangeAddress bMerge = merge(sh, r1, r3, 1, 1);
                        setCenter(sh, r1, 1, "т" + t, centerBorder);

                        // C: место — ТЕПЕРЬ с переносом строк
                        setText(getOrCreateRow(sh, r1), 2, place, leftWrapBorder);

                        // D: текст — уже со стилем с переносом
                        setText(getOrCreateRow(sh, r1), 3,
                                "Суммарные источники шума\n(работает лифтовое оборудование)",
                                centerWrapBorder);

                        // E..S: +/- (первая строка)
                        for (int i = 0; i <= (18 - 4); i++) {
                            setText(getOrCreateRow(sh, r1), 4 + i, PLUS_MINUS[i], centerBorder);
                        }

                        // T–V и W–Y (первая строка)
                        CellRangeAddress tv1 = merge(sh, r1, r1, 19, 21);
                        setCenter(sh, r1, 19, "-", centerBorder);
                        CellRangeAddress wy1 = merge(sh, r1, r1, 22, 24);
                        setCenter(sh, r1, 22, "-", centerBorder);

                        // Вторая строка: C–I «Поправка ...»
                        CellRangeAddress ci2 = merge(sh, r2, r2, 2, 8);
                        setText(getOrCreateRow(sh, r2), 2, "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r2), c, "-", centerBorder);

                        // T–V и W–Y (вторая строка)
                        CellRangeAddress tv2 = merge(sh, r2, r2, 19, 21);
                        setCenter(sh, r2, 19, "2", centerBorder);
                        CellRangeAddress wy2 = merge(sh, r2, r2, 22, 24);
                        setCenter(sh, r2, 22, "2", centerBorder);

                        // Третья строка: C–I «Уровни звука ...»
                        CellRangeAddress ci3 = merge(sh, r3, r3, 2, 8);
                        setText(getOrCreateRow(sh, r3), 2, "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r3), c, "-", centerBorder);

                        for (CellRangeAddress rg : new CellRangeAddress[]{aMerge, bMerge, tv1, wy1, ci2, tv2, wy2, ci3}) {
                            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rg, sh);
                        }

                        row += 3;
                    }
                }
            }
        }
    }
}
