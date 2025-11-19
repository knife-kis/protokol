package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import java.util.Map;

import static ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon.*;

public final class NoiseAutoSheetWriter {
    private NoiseAutoSheetWriter() {}

    /** Обёртка: начинает с 2-й строки (после простой шапки). */
    public static void appendAutoRoomBlocks(Workbook wb, Sheet sh,
                                            Building building,
                                            Map<String, DatabaseManager.NoiseValue> byKey) {
        appendAutoRoomBlocksFromRow(wb, sh, building, byKey, 2);
    }
    // СТАРАЯ СИГНАТУРА — для совместимости со старым вызовом обёртки
    public static void appendAutoRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                   Building building,
                                                   java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                   int startRow) {
        appendAutoRoomBlocksFromRow(
                wb, sh, building, byKey, startRow,
                java.util.Collections.emptyMap(),
                ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind.AUTO_DAY
        );
    }

    /** Наполняет лист «Авто» блоками т1/т2/т3. Только квартиры (APARTMENT), только где включён «Авто». */
    public static void appendAutoRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                   Building building,
                                                   Map<String, DatabaseManager.NoiseValue> byKey,
                                                   int startRow,
                                                   Map<String, double[]> thresholds,
                                                   ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind) {
        if (building == null) return;

        Font f8 = wb.createFont();
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

        // C — место измерений: перенос + по левому краю
        CellStyle leftWrapBorder = wb.createCellStyle();
        leftWrapBorder.cloneStyleFrom(centerBorder);
        leftWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftWrapBorder.setWrapText(true);

        // D — текст источника: перенос (центр)
        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerBorder);
        centerWrapBorder.setWrapText(true);

        CellStyle leftNoWrapBorder = wb.createCellStyle();
        leftNoWrapBorder.cloneStyleFrom(centerBorder);
        leftNoWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftNoWrapBorder.setWrapText(false);

        // Е..S: последовательность +/- (как на лифте/ИТО)
        final String[] PLUS_MINUS = { "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-" };

        int row = Math.max(0, startRow);
        int no  = 1;

        var sections = building.getSections();
        boolean multiSections = sections != null && sections.size() > 1;

        var floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor fl : floors) {
            String floorNum = (fl.getNumber() == null) ? "" : fl.getNumber().trim();

            var spaces = new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space sp : spaces) {
                if (sp == null || sp.getType() != Space.SpaceType.APARTMENT) continue; // только квартиры

                String spaceId = (sp.getIdentifier() == null) ? "" : sp.getIdentifier().trim();

                var rooms = new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room rm : rooms) {
                    String roomName = (rm.getName() == null) ? "" : rm.getName().trim();
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null || !nv.autoSrc) continue; // нужен источник «Авто»

                    Section sec = null;
                    if (multiSections && fl.getSectionIndex() >= 0 && fl.getSectionIndex() < sections.size()) {
                        sec = sections.get(fl.getSectionIndex());
                    }

                    String place = formatPlace(building, sec, fl, sp, rm);
                    String dText = "Суммарные источники шума (движение автотранспорта)";

                    for (int t = 1; t <= 3; t++) {
                        int r1 = row, r2 = row + 1, r3 = row + 2;

                        // сетка A..Y
                        for (int rr = r1; rr <= r3; rr++) {
                            Row cur = getOrCreateRow(sh, rr);
                            for (int c = 0; c <= 24; c++) {
                                Cell cell = getOrCreateCell(cur, c);
                                cell.setCellStyle(centerBorder);
                            }
                        }

                        // A №п/п (merge) и B т1/т2/т3 (merge)
                        CellRangeAddress aMerge = merge(sh, r1, r3, 0, 0);
                        setCenter(sh, r1, 0, String.valueOf(no++), centerBorder);

                        CellRangeAddress bMerge = merge(sh, r1, r3, 1, 1);
                        setCenter(sh, r1, 1, "т" + t, centerBorder);

                        // C — место (wrap), D — текст источника (wrap)
                        setText(getOrCreateRow(sh, r1), 2, place, leftWrapBorder);
                        setText(getOrCreateRow(sh, r1), 3, dText, centerWrapBorder);

                        // E..S: +/-; T–V и W–Y — объединения
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

                        // 2-я строка блока
                        CellRangeAddress ci2 = merge(sh, r2, r2, 2, 8);
                        setText(getOrCreateRow(sh, r2), 2, "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r2), c, "-", centerBorder);
                        CellRangeAddress tv2 = merge(sh, r2, r2, 19, 21);
                        setCenter(sh, r2, 19, "2", centerBorder);
                        CellRangeAddress wy2 = merge(sh, r2, r2, 22, 24);
                        setCenter(sh, r2, 22, "2", centerBorder);

                        // 3-я строка блока
                        CellRangeAddress ci3 = merge(sh, r3, r3, 2, 8);
                        setText(getOrCreateRow(sh, r3), 2, "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ)", leftNoWrapBorder);
                        for (int c = 9; c <= 18; c++) setText(getOrCreateRow(sh, r3), c, "-", centerBorder);

                        fillNoiseThirdRowDiffs(sh, r1, r2, r3, thirdRowDiff, thirdRowFixed, plusMinusNoLR);

                        // рамки объединений
                        for (CellRangeAddress rg : new CellRangeAddress[]{aMerge, bMerge, tv1, wy1, ci2, tv2, wy2, ci3}) {
                            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rg, sh);
                            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rg, sh);
                        }

                        // авто-подгонка высоты первой строки блока по C и D
                        adjustRowHeightForWrapped(sh, r1, 0.53, new int[]{2, 3}, new String[]{place, dText});
                        // фикс-высоты 2 и 3 строк
                        setRowHeightCm(sh, r2, 0.53);
                        setRowHeightCm(sh, r3, 0.53);

                        row += 3;
                    }
                }
            }
        }
    }
}
