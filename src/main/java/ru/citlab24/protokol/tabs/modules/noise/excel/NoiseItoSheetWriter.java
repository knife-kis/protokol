package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import static ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon.*;

/**
 * Наполнение листов ИТО:
 * - «шум неж ИТО»    — помещения OFFICE и PUBLIC_SPACE, если выбран ≥1 ИТО-источник (Вент/Завеса/ИТП/ПНС/Э/Щ)
 * - «шум жил ИТО день/ночь» — только APARTMENT, те же правила источников
 * - «шум зум» — только APARTMENT и только «зачистное устройство мусоропровода» (добавляется к ИТО жилые день)
 *
 * Сетка и стили как на «лифт ночь» (упрощённая шапка), колонка D — динамический текст по выбранным источникам.
 */
public final class NoiseItoSheetWriter {

    private NoiseItoSheetWriter() {}

    /** Нежилые ИТО: берём OFFICE и PUBLIC_SPACE. Стартовая строка передаётся явно (обычно 2). */
    public static int appendItoNonResRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                       Building building,
                                                       java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                       int startRow,
                                                       int startNo,
                                                       java.util.Map<String, double[]> thresholds,
                                                       ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind,
                                                       boolean addNormativeRow) {
        return appendItoRoomBlocksFromRow(wb, sh, building, byKey, startRow, startNo, thresholds, sheetKind,
                sp -> sp != null && (sp.getType() == Space.SpaceType.OFFICE || sp.getType() == Space.SpaceType.PUBLIC_SPACE),
                NoiseItoSheetWriter::hasItoSources,
                NoiseItoSheetWriter::itoDText,
                addNormativeRow);
    }

    /** Жилые ИТО (день/ночь): только APARTMENT. Стартовая строка передаётся явно (обычно 2). */
    public static int appendItoResRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                    Building building,
                                                    java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                    int startRow,
                                                    int startNo,
                                                    java.util.Map<String, double[]> thresholds,
                                                    ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind,
                                                    boolean addNormativeRow) {
        return appendItoRoomBlocksFromRow(wb, sh, building, byKey, startRow, startNo, thresholds, sheetKind,
                sp -> sp != null && sp.getType() == Space.SpaceType.APARTMENT,
                NoiseItoSheetWriter::hasItoSources,
                NoiseItoSheetWriter::itoDText,
                addNormativeRow);
    }

    /** ЗУМ (жилые): только APARTMENT, добавляется к «шум жил ИТО день». */
    public static int appendZumResRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                    Building building,
                                                    java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                    int startRow,
                                                    int startNo,
                                                    java.util.Map<String, double[]> thresholds,
                                                    ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind,
                                                    boolean addNormativeRow) {
        return appendItoRoomBlocksFromRow(wb, sh, building, byKey, startRow, startNo, thresholds, sheetKind,
                sp -> sp != null && sp.getType() == Space.SpaceType.APARTMENT,
                NoiseItoSheetWriter::hasZumSource,
                NoiseItoSheetWriter::zumDText,
                addNormativeRow);
    }

    /* ===== Общая реализация для ИТО ===== */
    private static int appendItoRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                  Building building,
                                                  java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                  int startRow,
                                                  int startNo,
                                                  java.util.Map<String, double[]> thresholds,
                                                  ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind,
                                                  java.util.function.Predicate<Space> spacePredicate,
                                                  java.util.function.Predicate<DatabaseManager.NoiseValue> sourcesPredicate,
                                                  java.util.function.Function<DatabaseManager.NoiseValue, String> dTextProvider,
                                                  boolean addNormativeRow) {
        if (building == null) return startNo;

        org.apache.poi.ss.usermodel.Font f8 = wb.createFont();
        f8.setFontName("Arial");
        f8.setFontHeightInPoints((short)8);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.setAlignment(HorizontalAlignment.CENTER);
        centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
        centerBorder.setWrapText(false);
        centerBorder.setFont(f8);
        setThinBorder(centerBorder);

        CellStyle plusMinusNoLR = wb.createCellStyle();
        plusMinusNoLR.cloneStyleFrom(centerBorder);
        plusMinusNoLR.setBorderLeft(BorderStyle.NONE);
        plusMinusNoLR.setBorderRight(BorderStyle.NONE);

        short oneDecimalFormat = wb.createDataFormat().getFormat("0.0");

        CellStyle thirdRowDiff = wb.createCellStyle();
        thirdRowDiff.cloneStyleFrom(centerBorder);
        thirdRowDiff.setBorderRight(BorderStyle.NONE);
        thirdRowDiff.setDataFormat(oneDecimalFormat);

        CellStyle thirdRowFixed = wb.createCellStyle();
        thirdRowFixed.cloneStyleFrom(centerBorder);
        thirdRowFixed.setBorderLeft(BorderStyle.NONE);

        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerBorder);
        centerWrapBorder.setWrapText(true); // D с переносом

        CellStyle leftWrapBorder = wb.createCellStyle();
        leftWrapBorder.cloneStyleFrom(centerBorder);
        leftWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftWrapBorder.setWrapText(true);   // C с переносом

        CellStyle leftNoWrapBorder = wb.createCellStyle();
        leftNoWrapBorder.cloneStyleFrom(centerBorder);
        leftNoWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftNoWrapBorder.setWrapText(false);

        final String[] PLUS_MINUS = { "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-" };

        int row = Math.max(0, startRow);
        int no  = Math.max(1, startNo);
        NoiseSheetCommon.PageBreakTracker pageTracker = NoiseSheetCommon.createPageBreakTracker(sh, row);

        java.util.List<Section> sections = building.getSections();
        boolean multiSections = sections != null && sections.size() > 1;

        java.util.List<Floor> floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor fl : floors) {
            String floorNum = (fl.getNumber() == null) ? "" : fl.getNumber().trim();

            java.util.List<Space> spaces = new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space sp : spaces) {
                if (!spacePredicate.test(sp)) continue;

                String spaceId = (sp.getIdentifier() == null) ? "" : sp.getIdentifier().trim();

                java.util.List<Room> rooms = new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room rm : rooms) {
                    String roomName = (rm.getName() == null) ? "" : rm.getName().trim();
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (!sourcesPredicate.test(nv)) continue;

                    Section sec = null;
                    if (multiSections && fl.getSectionIndex() >= 0 && fl.getSectionIndex() < sections.size()) {
                        sec = sections.get(fl.getSectionIndex());
                    }

                    String place = formatPlace(building, sec, fl, sp, rm);
                    String dText = dTextProvider.apply(nv);

                    for (int t = 1; t <= 3; t++) {
                        int r1 = row, r2 = row + 1, r3 = row + 2;

                        // базовая сетка A..Y
                        for (int rr = r1; rr <= r3; rr++) {
                            Row cur = getOrCreateRow(sh, rr);
                            for (int c = 0; c <= 24; c++) {
                                Cell cell = getOrCreateCell(cur, c);
                                cell.setCellStyle(centerBorder);
                            }
                        }

                        CellRangeAddress aMerge = merge(sh, r1, r3, 0, 0);
                        setCenter(sh, r1, 0, String.valueOf(no++), centerBorder);
                        CellRangeAddress bMerge = merge(sh, r1, r3, 1, 1);
                        setCenter(sh, r1, 1, "т" + t, centerBorder);

                        // C — место (wrap)
                        setText(getOrCreateRow(sh, r1), 2, place, leftWrapBorder);

                        // D — динамический текст (wrap)
                        setText(getOrCreateRow(sh, r1), 3, dText, centerWrapBorder);

                        // E..S: +/- (первая строка)
                        for (int i = 0; i <= (18 - 4); i++) {
                            setText(getOrCreateRow(sh, r1), 4 + i, PLUS_MINUS[i], centerBorder);
                        }

                        CellRangeAddress tv1 = merge(sh, r1, r1, 19, 21);
                        setCenter(sh, r1, 19, "-", centerBorder);
                        CellRangeAddress wy1 = merge(sh, r1, r1, 22, 24);
                        setCenter(sh, r1, 22, "-", centerBorder);

                        // <<< ВСТАВКА ГЕНЕРАЦИИ ПО ПОРОГАМ ДЛЯ t1 >>>
                        ru.citlab24.protokol.tabs.modules.noise.NoiseExcelExporter
                                .fillEqMaxFirstRow(sh, r1, sheetKind, thresholds);

                        // Вторая строка
                        CellRangeAddress ci2 = merge(sh, r2, r2, 2, 8);
                        setText(getOrCreateRow(sh, r2), 2, "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r2), c, "-", centerBorder);
                        CellRangeAddress tv2 = merge(sh, r2, r2, 19, 21);
                        setCenter(sh, r2, 19, "2", centerBorder);
                        CellRangeAddress wy2 = merge(sh, r2, r2, 22, 24);
                        setCenter(sh, r2, 22, "2", centerBorder);

                        // Третья строка
                        CellRangeAddress ci3 = merge(sh, r3, r3, 2, 8);
                        setText(getOrCreateRow(sh, r3), 2, "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r3), c, "-", centerBorder);

                        fillNoiseThirdRowDiffs(sh, r1, r2, r3, thirdRowDiff, thirdRowFixed, plusMinusNoLR);

                        for (CellRangeAddress rg : new CellRangeAddress[]{aMerge, bMerge, tv1, wy1, ci2, tv2, wy2, ci3}) {
                            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rg, sh);
                        }

                        // Явная подгонка высоты первой строки по C и D
                        adjustRowHeightForWrapped(sh, r1, 0.53, new int[]{2, 3}, new String[]{place, dText});

                        // фикс-высоты для 2 и 3 строк
                        setRowHeightCm(sh, r2, 0.53);
                        setRowHeightCm(sh, r3, 0.53);

                        if (pageTracker != null) {
                            double blockHeight = NoiseSheetCommon.blockHeightPoints(sh, r1, 3);
                            pageTracker.keepBlockTogether(sh, r1, blockHeight);
                        }

                        row += 3;
                    }
                }
            }
        }

        if (addNormativeRow) {
            NormativeRow norm = normativeFor(sheetKind);
            if (norm != null) {
                NoiseSheetCommon.appendNormativeRow(wb, sh, row, norm.text, norm.eq, norm.max);
            }
        }

        return no;
    }

    /* ===== Правила ИТО ===== */

    /** Есть ли выбранные источники для ИТО (без «Лифт»). */
    private static boolean hasItoSources(DatabaseManager.NoiseValue nv) {
        return nv != null && (nv.vent || nv.heatCurtain || nv.itp || nv.pns || nv.electrical);
    }

    /** Текст для колонки D на листах ИТО. */
    private static String itoDText(DatabaseManager.NoiseValue nv) {
        java.util.List<String> parts = new java.util.ArrayList<>(6);
        if (nv.vent)        parts.add("вентиляции");
        if (nv.heatCurtain) parts.add("тепловой завесы");
        if (nv.itp)         parts.add("ИТП");
        if (nv.pns)         parts.add("ПНС");
        if (nv.electrical)  parts.add("электрощитовой");
        String joined = String.join(", ", parts);
        return "Суммарные источники шума (работает оборудование " + joined + ")";
    }

    private static boolean hasZumSource(DatabaseManager.NoiseValue nv) {
        return nv != null && nv.zum;
    }

    private static String zumDText(DatabaseManager.NoiseValue nv) {
        return "Суммарные источники шума (работает оборудование зачистного устройства мусоропровода)";
    }

    private static NormativeRow normativeFor(ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind kind) {
        if (kind == null) return null;
        switch (kind) {
            case ITO_NONRES:
                return new NormativeRow(NoiseSheetCommon.NORM_SP_51, "45", "60");
            case ITO_RES_DAY:
                return new NormativeRow(NoiseSheetCommon.NORM_SANPIN_DAY, "35", "50");
            case ITO_RES_NIGHT:
                return new NormativeRow(NoiseSheetCommon.NORM_SANPIN_NIGHT, "25", "40");
            case ZUM_DAY:
                return new NormativeRow(NoiseSheetCommon.NORM_SANPIN_DAY, "35", "50");
            default:
                return null;
        }
    }

    private static final class NormativeRow {
        final String text;
        final String eq;
        final String max;

        NormativeRow(String text, String eq, String max) {
            this.text = text;
            this.eq = eq;
            this.max = max;
        }
    }
}
