package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.*;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

public final class NoiseExcelExporter {

    private NoiseExcelExporter() {}

    /** Экспорт «Шумы/Лифт»: создаёт листы «шум лифт день» и «шум лифт ночь». */
    /** Экспорт «Шумы/Лифт».
     *  @param dateLine строка «Дата, время проведения измерений …» для A7–Y7 на листе «шум лифт день»
     */
    public static void exportLift(Building building,
                                  Map<String, DatabaseManager.NoiseValue> byKey,
                                  Component parent,
                                  java.util.Map<NoiseTestKind, String> dateLines) {
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet day        = wb.createSheet("шум лифт день");
            Sheet night      = wb.createSheet("шум лифт ночь");
            Sheet nonResIto  = wb.createSheet("шим неж ИТО");
            Sheet resIto     = wb.createSheet("шум жил ИТО");
            Sheet autoDay    = wb.createSheet("шум авто день");
            Sheet autoNight  = wb.createSheet("шум авто ночь");
            Sheet site       = wb.createSheet("шум площадка");

            // Единые ширины + печать в 1 страницу по ширине
            setupPage(day);       setupColumns(day);       shrinkColumnsToOnePrintedPage(day);
            setupPage(night);     setupColumns(night);     shrinkColumnsToOnePrintedPage(night);
            setupPage(nonResIto); setupColumns(nonResIto); shrinkColumnsToOnePrintedPage(nonResIto);
            setupPage(resIto);    setupColumns(resIto);    shrinkColumnsToOnePrintedPage(resIto);
            setupPage(autoDay);   setupColumns(autoDay);   shrinkColumnsToOnePrintedPage(autoDay);
            setupPage(autoNight); setupColumns(autoNight); shrinkColumnsToOnePrintedPage(autoNight);
            setupPage(site);      setupColumns(site);      shrinkColumnsToOnePrintedPage(site);

            // «Лифт день»: полная шапка + блоки т1/т2/т3 как раньше
            writeLiftHeader(wb, day, safeDateLine(dateLines, NoiseTestKind.LIFT_DAY));
            appendLiftRoomBlocks(wb, day, building, byKey); // внутри начинает с 7-й строки

            // «Лифт ночь»: упрощённая 2-строчная шапка + ТЕ ЖЕ блоки т1/т2/т3, старт с 3-й строки
            writeSimpleHeader(wb, night, safeDateLine(dateLines, NoiseTestKind.LIFT_NIGHT));
            appendLiftRoomBlocksFromRow(wb, night, building, byKey, 2);

            // Остальные листы — пока только две верхние строки (нумерация + своё время)
            writeSimpleHeader(wb, nonResIto, safeDateLine(dateLines, NoiseTestKind.ITO_NONRES));
            writeSimpleHeader(wb, resIto,    safeDateLine(dateLines, NoiseTestKind.ITO_RES));
            writeSimpleHeader(wb, autoDay,   safeDateLine(dateLines, NoiseTestKind.AUTO_DAY));
            writeSimpleHeader(wb, autoNight, safeDateLine(dateLines, NoiseTestKind.AUTO_NIGHT));
            // site — оставляем пустым

            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Сохранить Excel (шумы)");
            fc.setSelectedFile(new File("шумы.xlsx"));
            if (fc.showSaveDialog(parent) == JFileChooser.APPROVE_OPTION) {
                File file = ensureXlsx(fc.getSelectedFile());
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

    /** Заполняет блоки данных (по 3 строки) для всех комнат, где выбран источник «Лифт».
     *  Возвращает индекс следующей свободной строки.
     */
    private static int writeLiftData(Workbook wb, Sheet sh, int startRow,
                                     Building building,
                                     Map<String, DatabaseManager.NoiseValue> byKey) {

        // ===== шрифты/стили 8 пт =====
        org.apache.poi.ss.usermodel.Font f8 = wb.createFont();
        f8.setFontName("Arial");
        f8.setFontHeightInPoints((short)8);

        CellStyle center = wb.createCellStyle();
        center.setAlignment(HorizontalAlignment.CENTER);
        center.setVerticalAlignment(VerticalAlignment.CENTER);
        center.setWrapText(false);
        center.setFont(f8);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.cloneStyleFrom(center);
        setThinBorder(centerBorder);

        CellStyle leftNoWrap = wb.createCellStyle();
        leftNoWrap.cloneStyleFrom(center);
        leftNoWrap.setAlignment(HorizontalAlignment.LEFT);

        CellStyle leftNoWrapBorder = wb.createCellStyle();
        leftNoWrapBorder.cloneStyleFrom(leftNoWrap);
        setThinBorder(leftNoWrapBorder);

        CellStyle centerWrapBorder = wb.createCellStyle();
        centerWrapBorder.cloneStyleFrom(centerBorder);
        centerWrapBorder.setWrapText(true);

        // Плюсы/минусы для 1-й строки блока: E..Y (21 колонка, E=4..Y=24)
        final String[] PLUS_MINUS = {
                "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-","-"
        };

        int row = startRow;
        int pointIdx = 1; // № п/п (колонка A)

        // Быстрая навигация по секциям
        java.util.List<Section> sections = building.getSections();
        // Отсортируем этажи/помещения/комнаты по position (как в UI)
        java.util.List<Floor> floors = new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(Floor::getPosition));

        for (Floor f : floors) {
            int secIdx = Math.max(0, f.getSectionIndex());
            Section sec = (secIdx >= 0 && secIdx < sections.size()) ? sections.get(secIdx) : null;

            java.util.List<Space> spaces = new java.util.ArrayList<>(f.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(Space::getPosition));

            for (Space s : spaces) {
                java.util.List<Room> rooms = new java.util.ArrayList<>(s.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(Room::getPosition));

                for (Room r : rooms) {
                    // Ключ, как в NoiseTab/saveSelectionsByKey()
                    String floorNum = (f.getNumber() == null) ? "" : f.getNumber().trim();
                    String spaceId  = (s.getIdentifier() == null) ? "" : s.getIdentifier().trim();
                    String roomName = (r.getName() == null) ? "" : r.getName().trim();
                    String key = secIdx + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null || !nv.lift) continue; // только если включён «Лифт»

                    // -------- три строки: 1.59см, 0.53см, 0.53см --------
                    int r1 = row, r2 = row + 1, r3 = row + 2;

                    setRowHeightCm(sh, r1, 1.59);
                    setRowHeightCm(sh, r2, 0.53);
                    setRowHeightCm(sh, r3, 0.53);

                    Row R1 = getOrCreateRow(sh, r1);
                    Row R2 = getOrCreateRow(sh, r2);
                    Row R3 = getOrCreateRow(sh, r3);

                    // Границы в каждой ячейке блока A..Y (чтобы сетка была везде)
                    for (int rr = r1; rr <= r3; rr++) {
                        Row cur = getOrCreateRow(sh, rr);
                        for (int c = 0; c <= 24; c++) {
                            Cell cell = getOrCreateCell(cur, c);
                            if (cell.getCellStyle() == null || cell.getCellStyle() == center)
                                cell.setCellStyle(centerBorder);
                        }
                    }

                    // A: объединяем на 3 строки, пишем № п/п
                    CellRangeAddress aMerge = merge(sh, r1, r3, 0, 0);
                    setText(R1, 0, String.valueOf(pointIdx++), centerBorder);
                    // рамка по периметру объединения
                    org.apache.poi.ss.util.RegionUtil.setBorderTop   (BorderStyle.THIN, aMerge, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, aMerge, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderLeft  (BorderStyle.THIN, aMerge, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderRight (BorderStyle.THIN, aMerge, sh);

                    // B: т1, т2, т3 (НЕ объединяем — по строке)
                    setText(R1, 1, "т1", centerBorder);
                    setText(R2, 1, "т2", centerBorder);
                    setText(R3, 1, "т3", centerBorder);

                    // C (1-я строка): место измерений — секция (если >1), этаж/помещение/комната
                    String place = formatPlace(building, sec, f, s, r);
                    setText(R1, 2, place, leftNoWrapBorder);

                    // D (1-я строка): «Суммарные источники шума (работает лифтовое оборудование)»
                    setText(R1, 3, "Суммарные источники шума (работает лифтовое оборудование)", leftNoWrapBorder);

                    // E..Y (1-я строка): последовательность + - - + - - ...
                    for (int i = 0; i < PLUS_MINUS.length; i++) {
                        setText(R1, 4 + i, PLUS_MINUS[i], centerBorder);
                    }

                    // 2-я строка: объединяем C–I, текст «Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)»
                    CellRangeAddress ci2 = merge(sh, r2, r2, 2, 8);
                    setText(R2, 2, "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)", leftNoWrapBorder);
                    org.apache.poi.ss.util.RegionUtil.setBorderTop   (BorderStyle.THIN, ci2, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, ci2, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderLeft  (BorderStyle.THIN, ci2, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderRight (BorderStyle.THIN, ci2, sh);

                    // J..S = "-" ; T="2"; U="" ; V="2"; W..Y=""
                    for (int c = 9; c <= 18; c++) setText(R2, c, "-", centerBorder); // J..S
                    setText(R2, 19, "2", centerBorder); // T
                    setText(R2, 20, "", centerBorder);  // U (узкая)
                    setText(R2, 21, "2", centerBorder); // V
                    // W..Y пустые (рамки уже заданы выше)

                    // 3-я строка: объединяем C–I, текст «Уровни звука ... с учетом поправок, дБА (дБ)»
                    CellRangeAddress ci3 = merge(sh, r3, r3, 2, 8);
                    setText(R3, 2, "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ)", leftNoWrapBorder);
                    org.apache.poi.ss.util.RegionUtil.setBorderTop   (BorderStyle.THIN, ci3, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, ci3, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderLeft  (BorderStyle.THIN, ci3, sh);
                    org.apache.poi.ss.util.RegionUtil.setBorderRight (BorderStyle.THIN, ci3, sh);

                    // J..S = "-"; T..Y пусто
                    for (int c = 9; c <= 18; c++) setText(R3, c, "-", centerBorder);

                    // К следующему блоку
                    row += 3;
                }
            }
        }
        return row;
    }

    /** Формирование подписи «место измерений» для колонки C (1-я строка блока). */
    /** Колонка C (1-я строка блока): "<идентификатор помещения>, <комната>[, <название секции>]".
     *  Без указания этажа.
     */
    private static String formatPlace(Building b, Section sec, Floor f, Space s, Room r) {
        StringBuilder sb = new StringBuilder();

        String spaceId  = (s != null && s.getIdentifier() != null) ? s.getIdentifier().trim() : "";
        String roomName = (r != null && r.getName() != null)       ? r.getName().trim()       : "";

        if (!spaceId.isEmpty()) {
            sb.append(spaceId);
            if (!roomName.isEmpty()) sb.append(", ");
        }
        sb.append(roomName);

        boolean multiSections = b != null && b.getSections() != null && b.getSections().size() > 1;
        String secName = (sec != null && sec.getName() != null) ? sec.getName().trim() : "";
        if (multiSections && !secName.isEmpty()) {
            sb.append(", ").append(secName); // просто название секции, без слова "секция" и без скобок
        }
        return sb.toString();
    }

    /* ===================== ВНУТРЕННЕЕ ===================== */

    private static File ensureXlsx(File f) {
        String name = f.getName().toLowerCase();
        if (!name.endsWith(".xlsx")) {
            return new File(f.getParentFile(), f.getName() + ".xlsx");
        }
        return f;
    }

    private static void setupPage(Sheet sh) {
        PrintSetup ps = sh.getPrintSetup();

        // Бумага и ориентация
        ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
        ps.setLandscape(true);

        // Поля указываются в дюймах (оставляем твои значения)
        sh.setMargin(Sheet.LeftMargin, 1.80 / 2.54);
        sh.setMargin(Sheet.RightMargin, 1.48 / 2.54);

        // Печать: умещать по ШИРИНЕ в одну страницу, высота — сколько потребуется
        sh.setAutobreaks(true);
        sh.setFitToPage(true);
        ps.setFitWidth((short) 1);
        ps.setFitHeight((short) 0); // 0 = сколько нужно по высоте

        // По желанию можно центрировать по горизонтали (не обязательно)
        // sh.setHorizontallyCenter(true);
    }


    private static void setupColumns(Sheet sh) {
        // Помощник: приблизительно переводим см -> ширина колонки (в "символах" * 256)
        // 1 см ≈ 37.8 px; 1 символ (Calibri 11) ≈ 7 px → ~ cm * (37.8/7) = cm * 5.4 символа
        java.util.function.BiConsumer<Integer, Double> setCm = (col, cm) -> {
            double chars = cm * 5.4;
            sh.setColumnWidth(col, (int) Math.round(chars * 256));
        };

        // A:0.82  B:0.95  C:4.55  D:3.70  E–I:0.69
        setCm.accept(0, 0.82); // A
        setCm.accept(1, 0.95); // B
        setCm.accept(2, 4.55); // C
        setCm.accept(3, 3.70); // D
        for (int c = 4; c <= 8; c++) setCm.accept(c, 0.69); // E..I

        // J–R:0.95  S:0.90  T:0.71  U:0.32  V:0.58  W:0.71  X:0.32  Y:0.64
        for (int c = 9; c <= 17; c++) setCm.accept(c, 0.95); // J..R
        setCm.accept(18, 0.90); // S
        setCm.accept(19, 0.71); // T
        setCm.accept(20, 0.32); // U
        setCm.accept(21, 0.58); // V
        setCm.accept(22, 0.71); // W
        setCm.accept(23, 0.32); // X
        setCm.accept(24, 0.64); // Y
    }

    /** Вся «шапка» по ТЗ. НИГДЕ не делаем bold. A1 и B1 — без переноса. */
    private static void writeLiftHeader(Workbook wb, Sheet sh, String dateLine) {
        // ===== Шрифты =====
        Font f8  = wb.createFont();  f8.setFontName("Arial");  f8.setFontHeightInPoints((short) 8);
        Font f10 = wb.createFont();  f10.setFontName("Arial"); f10.setFontHeightInPoints((short)10);
        Font f7  = wb.createFont();  f7.setFontName("Arial");  f7.setFontHeightInPoints((short) 7);

        // ===== Базовые стили =====
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

        // Вертикальные варианты
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

        // ===== Высоты строк =====
        sh.setDefaultRowHeightInPoints(cmToPt(0.53));
        setRowHeightCm(sh, 0, 0.53);
        setRowHeightCm(sh, 1, 0.53);
        setRowHeightCm(sh, 2, 0.53);
        setRowHeightCm(sh, 3, 0.74);
        setRowHeightCm(sh, 4, 2.94);
        setRowHeightCm(sh, 5, 0.53);

        // ===== A1 / A2 =====
        Row r1 = getOrCreateRow(sh, 0);
        Cell a1 = getOrCreateCell(r1, 0);
        a1.setCellValue("16. Результаты измерений виброакустических факторов");
        a1.setCellStyle(leftNoWrapF10);

        Row r2 = getOrCreateRow(sh, 1);
        Cell a2 = getOrCreateCell(r2, 0);
        a2.setCellValue("16.2. Шум:");
        a2.setCellStyle(leftNoWrapF10);

        // ===== Шапка 3–5 =====
        java.util.List<CellRangeAddress> merges = new java.util.ArrayList<>();

        // A3–A5
        merges.add(merge(sh, 2, 4, 0, 0));
        setCenter(sh, 2, 0, "№ п/п", verticalBorderF10);

        // B3–B5
        merges.add(merge(sh, 2, 4, 1, 1));
        setCenter(sh, 2, 1, "№ точки измерения", verticalBorderF10);

        // C3–C5
        merges.add(merge(sh, 2, 4, 2, 2));
        setCenter(sh, 2, 2, "Место измерений", centerBorderF10_noWrap);

        // D3–D5
        merges.add(merge(sh, 2, 4, 3, 3));
        setCenter(sh, 2, 3, "Источник шума\n(тип, вид, марка, условия замера)", centerWrapBorderF10);

        // E3–I3
        merges.add(merge(sh, 2, 2, 4, 8));
        setCenter(sh, 2, 4, "Характер шума", centerBorderF10_noWrap);

        // E4–F4
        merges.add(merge(sh, 3, 3, 4, 5));
        setCenter(sh, 3, 4, "по спектру", centerWrapBorderF7);

        // G4–I4
        merges.add(merge(sh, 3, 3, 6, 8));
        setCenter(sh, 3, 6, "по временным характеристикам", centerWrapBorderF7);

        // Ряд 5: E..I (вертикально)
        Row r5 = getOrCreateRow(sh, 4);
        setText(r5, 4, "широкополосный", verticalBorderF10);
        setText(r5, 5, "тональный",       verticalBorderF10);
        setText(r5, 6, "постоянный",      verticalBorderF10);
        setText(r5, 7, "непостоянный",    verticalBorderF10);
        setText(r5, 8, "импульсный",      verticalBorderF10);

        // J3–R4 заголовок
        merges.add(merge(sh, 2, 3, 9, 17));
        setCenter(sh, 2, 9,
                "Уровни звукового давления (дБ) ± U (дБ) в октавных полосах частот со среднегеометрическими частотами (Гц)",
                centerWrapBorder);

        // J5..R5 частоты (10 пт)
        String[] freqs = {"31,5","63","125","250","500","1000","2000","4000","8000"};
        for (int i = 0; i < freqs.length; i++) setText(r5, 9 + i, freqs[i], centerBorderF10);

        // S3–S5
        merges.add(merge(sh, 2, 4, 18, 18));
        setCenter(sh, 2, 18, "Уровни звука (дБА) ±U (дБ)", verticalWrapBorderF10);

        // T3–V5
        merges.add(merge(sh, 2, 4, 19, 21));
        setCenter(sh, 2, 19, "Эквивалентные уровни звука,  (дБА) ±U (дБ)", verticalWrapBorderF10);

        // W3–Y5
        merges.add(merge(sh, 2, 4, 22, 24));
        setCenter(sh, 2, 22, "Максимальные уровни звука  (дБА)", verticalWrapBorderF10);

        // Ряд 6: нумерация
        Row r6 = getOrCreateRow(sh, 5);
        for (int c = 0; c <= 18; c++) setText(r6, c, String.valueOf(c + 1), (c <= 3) ? centerBorderF10 : centerBorder);
        merges.add(merge(sh, 5, 5, 19, 21)); setCenter(sh, 5, 19, "20", centerBorder);
        merges.add(merge(sh, 5, 5, 22, 24)); setCenter(sh, 5, 22, "21", centerBorder);

        // Рамки по контурам объединений
        for (CellRangeAddress rgn : merges) {
            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rgn, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rgn, sh);
        }

        // ===== A7–Y7: по ЦЕНТРУ + рамка (ваше требование) =====
        if (dateLine != null && !dateLine.isBlank()) {
            setRowHeightCm(sh, 6, 0.53);                 // строка 7
            CellRangeAddress rng = merge(sh, 6, 6, 0, 24); // A7..Y7
            Row r7 = getOrCreateRow(sh, 6);
            Cell a7 = getOrCreateCell(r7, 0);
            a7.setCellValue(dateLine);

            CellStyle centerBorder8 = wb.createCellStyle();
            centerBorder8.cloneStyleFrom(centerBorder);
            centerBorder8.setFont(f8);
            centerBorder8.setAlignment(HorizontalAlignment.CENTER);
            centerBorder8.setVerticalAlignment(VerticalAlignment.CENTER);
            a7.setCellStyle(centerBorder8);

            // рамку по периметру слияния тоже ставим
            org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, rng, sh);
            org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, rng, sh);
        }
    }
    /**
     * Упрощённая шапка для листов: 1-я строка — как «строка 6» на дневном листе (нумерация),
     * 2-я строка — как A7–Y7 (дата/время) по центру в рамке.
     * Никакого bold, шрифт 8 пт, тонкие рамки как в основном листе.
     */
    private static void writeSimpleHeader(Workbook wb, Sheet sh, String dateLine) {
        // Шрифты/стили
        Font f8  = wb.createFont();  f8.setFontName("Arial");  f8.setFontHeightInPoints((short)8);

        CellStyle centerBorder = wb.createCellStyle();
        centerBorder.setAlignment(HorizontalAlignment.CENTER);
        centerBorder.setVerticalAlignment(VerticalAlignment.CENTER);
        centerBorder.setWrapText(false);
        centerBorder.setFont(f8);
        setThinBorder(centerBorder);

        // 1) Строка 1 — нумерация как на «шум лифт день», строка 6 (индекс 5).
        setRowHeightCm(sh, 0, 0.53);
        Row r1 = getOrCreateRow(sh, 0);
        for (int c = 0; c <= 18; c++) {
            setText(r1, c, String.valueOf(c + 1), centerBorder);
        }
        // T(19)..V(21) = "20" (merge)
        CellRangeAddress t20 = merge(sh, 0, 0, 19, 21);
        setCenter(sh, 0, 19, "20", centerBorder);
        // W(22)..Y(24) = "21" (merge)
        CellRangeAddress w21 = merge(sh, 0, 0, 22, 24);
        setCenter(sh, 0, 22, "21", centerBorder);
        // рамки
        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, t20, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, t20, sh);

        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, w21, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, w21, sh);

        // 2) Строка 2 — A..Y объединено, текст dateLine по центру + рамка
        setRowHeightCm(sh, 1, 0.53);
        CellRangeAddress a2y2 = merge(sh, 1, 1, 0, 24);
        Row r2 = getOrCreateRow(sh, 1);
        Cell a = getOrCreateCell(r2, 0);
        a.setCellValue(dateLine);
        a.setCellStyle(centerBorder);

        org.apache.poi.ss.util.RegionUtil.setBorderTop(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderBottom(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderLeft(BorderStyle.THIN, a2y2, sh);
        org.apache.poi.ss.util.RegionUtil.setBorderRight(BorderStyle.THIN, a2y2, sh);
    }

    /** Добавляет по всем комнатам блоки «3 строки × 3 замера (т1/т2/т3)» для лифтов. */
    private static void appendLiftRoomBlocks(Workbook wb, Sheet sh,
                                             Building building,
                                             Map<String, DatabaseManager.NoiseValue> byKey) {
        appendLiftRoomBlocksFromRow(wb, sh, building, byKey, 7);
    }
    /** То же, что appendLiftRoomBlocks, но с явной стартовой строкой (для листов без «A7–Y7» сверху). */
    private static void appendLiftRoomBlocksFromRow(Workbook wb, Sheet sh,
                                                    Building building,
                                                    Map<String, DatabaseManager.NoiseValue> byKey,
                                                    int startRow) {
        if (building == null) return;

        // ===== шрифты/стили 8 пт =====
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
        centerWrapBorder.setWrapText(true); // центр + перенос строки

        CellStyle leftNoWrapBorder = wb.createCellStyle();
        leftNoWrapBorder.cloneStyleFrom(centerBorder);
        leftNoWrapBorder.setAlignment(HorizontalAlignment.LEFT);
        leftNoWrapBorder.setWrapText(false);

        // Плюсы/минусы для 1-й строки блока: для E..S (E=4..S=18)
        final String[] PLUS_MINUS = { "+","-","-","+","-","-","-","-","-","-","-","-","-","-","-" };

        int row = Math.max(0, startRow);
        int no  = 1;  // № п/п

        java.util.List<ru.citlab24.protokol.tabs.models.Section> sections = building.getSections();
        boolean multiSections = sections != null && sections.size() > 1;

        java.util.List<ru.citlab24.protokol.tabs.models.Floor> floors =
                new java.util.ArrayList<>(building.getFloors());
        floors.sort(java.util.Comparator.comparingInt(ru.citlab24.protokol.tabs.models.Floor::getPosition));

        for (ru.citlab24.protokol.tabs.models.Floor fl : floors) {
            String floorNum = (fl.getNumber() == null) ? "" : fl.getNumber().trim();

            java.util.List<ru.citlab24.protokol.tabs.models.Space> spaces =
                    new java.util.ArrayList<>(fl.getSpaces());
            spaces.sort(java.util.Comparator.comparingInt(ru.citlab24.protokol.tabs.models.Space::getPosition));

            for (ru.citlab24.protokol.tabs.models.Space sp : spaces) {
                String spaceId = (sp.getIdentifier() == null) ? "" : sp.getIdentifier().trim();

                java.util.List<ru.citlab24.protokol.tabs.models.Room> rooms =
                        new java.util.ArrayList<>(sp.getRooms());
                rooms.sort(java.util.Comparator.comparingInt(ru.citlab24.protokol.tabs.models.Room::getPosition));

                for (ru.citlab24.protokol.tabs.models.Room rm : rooms) {
                    String roomName = (rm.getName() == null) ? "" : rm.getName().trim();
                    String key = Math.max(0, fl.getSectionIndex()) + "|" + floorNum + "|" + spaceId + "|" + roomName;

                    DatabaseManager.NoiseValue nv = byKey.get(key);
                    if (nv == null || !nv.lift) continue;

                    ru.citlab24.protokol.tabs.models.Section sec = null;
                    if (multiSections && fl.getSectionIndex() >= 0 && fl.getSectionIndex() < sections.size()) {
                        sec = sections.get(fl.getSectionIndex());
                    }

                    String place = formatPlace(building, sec, fl, sp, rm);

                    // т1/т2/т3 — ТРИ ОТДЕЛЬНЫХ БЛОКА, как на «дне»
                    for (int t = 1; t <= 3; t++) {
                        int r1 = row;
                        int r2 = row + 1;
                        int r3 = row + 2;

                        setRowHeightCm(sh, r1, 1.59);
                        setRowHeightCm(sh, r2, 0.53);
                        setRowHeightCm(sh, r3, 0.53);

                        // базовая сетка
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

                        // B: «т1/т2/т3» (merge 3 строки)
                        CellRangeAddress bMerge = merge(sh, r1, r3, 1, 1);
                        setCenter(sh, r1, 1, "т" + t, centerBorder);

                        // C (первая строка) — место
                        setCell(sh, r1, 2, place, leftNoWrapBorder);

                        // D (первая строка) — ЦЕНТР + перенос
                        setCell(sh, r1, 3, "Суммарные источники шума\n(работает лифтовое оборудование)", centerWrapBorder);

                        // E..S: +/- (первая строка)
                        for (int i = 0; i <= (18 - 4); i++) {
                            setCell(sh, r1, 4 + i, PLUS_MINUS[i], centerBorder);
                        }

                        // T–V и W–Y (первая строка) — "-"
                        CellRangeAddress tv1 = merge(sh, r1, r1, 19, 21);
                        setCenter(sh, r1, 19, "-", centerBorder);
                        CellRangeAddress wy1 = merge(sh, r1, r1, 22, 24);
                        setCenter(sh, r1, 22, "-", centerBorder);

                        // Вторая строка: C–I «Поправка ...»
                        CellRangeAddress ci2 = merge(sh, r2, r2, 2, 8);
                        setCell(sh, r2, 2, "Поправка (МИ Ш.13-2021 п.12.3.2.1.1) дБА (дБ)", centerBorder); // рамка + без переноса

                        // J..S = "-"
                        for (int c = 9; c <= 18; c++) setCell(sh, r2, c, "-", centerBorder);

                        // T–V и W–Y (вторая строка) — "2"
                        CellRangeAddress tv2 = merge(sh, r2, r2, 19, 21);
                        setCenter(sh, r2, 19, "2", centerBorder);
                        CellRangeAddress wy2 = merge(sh, r2, r2, 22, 24);
                        setCenter(sh, r2, 22, "2", centerBorder);

                        // Третья строка: C–I «Уровни звука ...»
                        CellRangeAddress ci3 = merge(sh, r3, r3, 2, 8);
                        setCell(sh, r3, 2, "Уровни звука (уровни звукового давления) с учетом поправок, дБА (дБ)", centerBorder);

                        // J..S = "-"
                        for (int c = 9; c <= 18; c++) setCell(sh, r3, c, "-", centerBorder);

                        // рамки для всех объединений
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

    /** Очищает тип этажа в названии: убирает «(жилой)», «(офисный)» и т.п. */
    private static String cleanFloorName(String s) {
        if (s == null) return "";
        // убираем любую скобочную приписку
        return s.replaceAll("\\s*\\(.*?\\)", "").trim();
    }

    // небольшие вспомогатели
    private static void setCell(Sheet sh, int r, int c, String text, CellStyle st) {
        Row row = getOrCreateRow(sh, r);
        Cell cell = getOrCreateCell(row, c);
        cell.setCellValue(text);
        cell.setCellStyle(st);
    }

    /* ===== helpers ===== */

    private static Row getOrCreateRow(Sheet sh, int r) {
        Row row = sh.getRow(r);
        return (row != null) ? row : sh.createRow(r);
    }

    private static Cell getOrCreateCell(Row row, int c) {
        Cell cell = row.getCell(c, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
        return (cell != null) ? cell : row.createCell(c);
    }

    private static CellRangeAddress merge(Sheet sh, int r1, int r2, int c1, int c2) {
        CellRangeAddress a = new CellRangeAddress(r1, r2, c1, c2);
        sh.addMergedRegion(a);
        return a;
    }

    private static void setCenter(Sheet sh, int r, int c, String text, CellStyle style) {
        Row row = getOrCreateRow(sh, r);
        Cell cell = getOrCreateCell(row, c);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }

    private static void setText(Row row, int c, String text, CellStyle style) {
        Cell cell = getOrCreateCell(row, c);
        cell.setCellValue(text);
        cell.setCellStyle(style);
    }
    private static void setThinBorder(CellStyle st) {
        st.setBorderTop(BorderStyle.THIN);
        st.setBorderBottom(BorderStyle.THIN);
        st.setBorderLeft(BorderStyle.THIN);
        st.setBorderRight(BorderStyle.THIN);
    }
    private static void setRowHeightCm(Sheet sh, int rowIndex, double cm) {
        getOrCreateRow(sh, rowIndex).setHeightInPoints(cmToPt(cm));
    }
    private static float cmToPt(double cm) {
        return (float) (cm * 72.0 / 2.54);
    }
    /**
     * Пропорционально уменьшает ширины столбцов так, чтобы сумма ширин (A..Y)
     * укладывалась в печатную область одной страницы по ширине (A4 landscape, текущие поля).
     * Работает в "колоночных единицах" Excel (256 = 1 символ).
     */
    private static void shrinkColumnsToOnePrintedPage(Sheet sh) {
        final int FIRST_COL = 0;   // A
        final int LAST_COL  = 24;  // Y

        // 1) Текущая суммарная ширина в единицах Excel (char*256)
        long sum = 0;
        for (int c = FIRST_COL; c <= LAST_COL; c++) {
            sum += sh.getColumnWidth(c);
        }
        if (sum <= 0) return;

        // 2) Целевая печатная ширина в символах (грубая, но стабильная оценка):
        //    ширина страницы (см) минус поля → перевод в "символы" (~5.4 символа на 1 см),
        //    затем в единицы Excel (×256).
        PrintSetup ps = sh.getPrintSetup();
        boolean landscape = ps.getLandscape();
        // A4: 21.0 × 29.7 см
        double pageWidthCm = landscape ? 29.7 : 21.0;

        // Поля заданы в дюймах → переводим в см
        double leftCm  = sh.getMargin(Sheet.LeftMargin)  * 2.54;
        double rightCm = sh.getMargin(Sheet.RightMargin) * 2.54;

        double printableCm = Math.max(0.1, pageWidthCm - (leftCm + rightCm)); // страховка от отрицательных
        double charsPerCm  = 5.4; // как ты уже использовал в setupColumns
        double targetChars = printableCm * charsPerCm;

        long targetUnits = (long) Math.floor(targetChars * 256.0);

        // 3) Если уже помещается — ничего не делаем
        if (sum <= targetUnits) return;

        // 4) Иначе считаем коэффициент и масштабирум каждую колонку
        double k = (double) targetUnits / (double) sum;

        // Ограничения, чтобы совсем узкие колонки не ушли в ноль
        final int MIN_UNITS = 1 * 256; // минимум 1 символ
        for (int c = FIRST_COL; c <= LAST_COL; c++) {
            int w = sh.getColumnWidth(c);
            int nw = (int) Math.max(MIN_UNITS, Math.round(w * k));
            sh.setColumnWidth(c, nw);
        }
    }
    private static String safeDateLine(java.util.Map<NoiseTestKind, String> map, NoiseTestKind k) {
        String s = (map != null) ? map.get(k) : null;
        return (s != null && !s.isBlank())
                ? s
                : "Дата, время проведения измерений __.__.____ c __:__ до __:__";
    }

}
