package ru.citlab24.protokol.tabs.modules.noise.excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import static ru.citlab24.protokol.tabs.modules.noise.excel.NoiseSheetCommon.*;

/**
 * Наполнение листа «шум площадка» (улица):
 * - включаем только этажи типа STREET и помещения OUTDOOR;
 * - берём комнаты, где выбран хотя бы один источник из двух: Авто (nv.autoSrc) или Поезд (временная мапа на nv.zum);
 * - каждая точка т1/т2/т3 — ОДНА строка (без трёхстрочного блока);
 * - C: только название комнаты (wrap);
 * - D: «Суммарные источники шума (движение …)»:
 *      авто → «… автотранспорта»; поезд → «… железнодорожного транспорта»; оба → «… авто- и железнодорожного транспорта»;
 * - E..S ставим шаблон +/- как в других листах;
 * - T..V и W..Y НЕ объединяем и оставляем пустыми.
 */
public final class NoiseSiteSheetWriter {

    private NoiseSiteSheetWriter() {}

    public static void appendSiteRoomRowsFromRow(Workbook wb, Sheet sh,
                                                 Building building,
                                                 java.util.Map<String, DatabaseManager.NoiseValue> byKey,
                                                 int startRow,
                                                 java.util.Map<String, double[]> thresholds,
                                                 ru.citlab24.protokol.tabs.modules.noise.NoiseTestKind sheetKind) {
        if (building == null) return;

        Font f8 = wb.createFont();
        f8.setFontName("Arial");
        f8.setFontHeightInPoints((short) 8);

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

        CellStyle noRightOneDecimal = wb.createCellStyle();
        noRightOneDecimal.cloneStyleFrom(centerBorder);
        noRightOneDecimal.setBorderRight(BorderStyle.NONE);
        noRightOneDecimal.setDataFormat(oneDecimalFormat);

        CellStyle noLeftBorder = wb.createCellStyle();
        noLeftBorder.cloneStyleFrom(centerBorder);
        noLeftBorder.setBorderLeft(BorderStyle.NONE);

        CellStyle leftWrapBorder = wb.createCellStyle();
        leftWrapBorder.cloneStyleFrom(centerBorder);
        leftWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftWrapBorder.setWrapText(true); // C — перенос

        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerBorder);
        centerWrapBorder.setWrapText(true); // D — перенос

        final String[] PLUS_MINUS = { "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-" };

        int row = Math.max(0, startRow);
        int no  = 1;

        java.util.List<Floor> floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor fl : floors) {
            if (fl.getType() != Floor.FloorType.STREET) continue;

            java.util.List<Space> spaces = new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space sp : spaces) {
                if (sp.getType() != Space.SpaceType.OUTDOOR) continue;

                String floorNum = ns(fl.getNumber());
                String spaceId  = ns(sp.getIdentifier());

                java.util.List<Room> rooms = new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room rm : rooms) {
                    String roomName = ns(rm.getName());
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    boolean auto  = nv != null && nv.autoSrc;
                    boolean train = nv != null && nv.zum; // «Поезд» временно храним в nv.zum
                    if (!auto && !train) continue;

                    String dText = siteDText(auto, train);

                    for (int t = 1; t <= 3; t++) {
                        int r = row++;

                        // базовая сетка A..Y — без объединений
                        Row rr = getOrCreateRow(sh, r);
                        for (int c = 0; c <= 24; c++) {
                            Cell cell = getOrCreateCell(rr, c);
                            cell.setCellStyle(centerBorder);
                            if (c >= 19) cell.setBlank(); // T..Y пустые по умолчанию
                        }
                        getOrCreateCell(rr, 19).setCellStyle(noRightOneDecimal);
                        getOrCreateCell(rr, 21).setCellStyle(noLeftBorder);
                        getOrCreateCell(rr, 22).setCellStyle(noRightOneDecimal);
                        getOrCreateCell(rr, 24).setCellStyle(noLeftBorder);

                        // A — № п/п
                        setCenter(sh, r, 0, String.valueOf(no++), centerBorder);
                        // B — т1/т2/т3
                        setCenter(sh, r, 1, "т" + t, centerBorder);
                        // C — только название комнаты (wrap)
                        setText(rr, 2, roomName, leftWrapBorder);
                        // D — динамический текст (wrap)
                        setText(rr, 3, dText, centerWrapBorder);

                        // E..S — плюс/минус
                        for (int i = 0; i <= (18 - 4); i++) {
                            setText(rr, 4 + i, PLUS_MINUS[i], centerBorder);
                        }

                        // <<< ВСТАВКА ГЕНЕРАЦИИ ПО ПОРОГАМ ДЛЯ ЭТОЙ СТРОКИ >>>
                        ru.citlab24.protokol.tabs.modules.noise.NoiseExcelExporter
                                .fillEqMaxFirstRow(sh, r, sheetKind, thresholds);

                        Cell uCell = getOrCreateCell(rr, 20);
                        uCell.setCellValue("±");
                        uCell.setCellStyle(plusMinusNoLR);
                        setText(rr, 21, "2,3", noLeftBorder);

                        Cell xCell = getOrCreateCell(rr, 23);
                        xCell.setCellValue("±");
                        xCell.setCellStyle(plusMinusNoLR);
                        setText(rr, 24, "2,3", noLeftBorder);

                        // высота строки по C и D
                        adjustRowHeightForWrapped(sh, r, 0.53, new int[]{2,3}, new String[]{roomName, dText});
                    }
                }
            }
        }

        NoiseSheetCommon.appendNormativeRow(wb, sh, row,
                NoiseSheetCommon.NORM_SANPIN, "45", "60");
    }

    private static String siteDText(boolean auto, boolean train) {
        if (auto && train) return "Суммарные источники шума (движение авто- и железнодорожного транспорта)";
        if (auto)         return "Суммарные источники шума (движение автотранспорта)";
        /*train*/         return "Суммарные источники шума (движение железнодорожного транспорта)";
    }

    private static String ns(String s) { return (s == null) ? "" : s.trim(); }
}
