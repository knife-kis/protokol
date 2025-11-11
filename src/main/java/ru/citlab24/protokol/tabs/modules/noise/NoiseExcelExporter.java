package ru.citlab24.protokol.tabs.modules.noise;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import ru.citlab24.protokol.db.DatabaseManager;
import ru.citlab24.protokol.tabs.models.Building;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;

public final class NoiseExcelExporter {

    private NoiseExcelExporter() {}

    /** Экспорт «Шумы/Лифт»: создаёт листы «шум лифт день» и «шум лифт ночь». */
    public static void exportLift(Building building,
                                  Map<String, DatabaseManager.NoiseValue> byKey,
                                  Component parent) {
        try (Workbook wb = new XSSFWorkbook()) {
            // Основные листы
            Sheet day        = wb.createSheet("шум лифт день");
            Sheet night      = wb.createSheet("шум лифт ночь");   // пустой (по ТЗ пока не заполняем)

            // Дополнительные листы
            Sheet nonResIto  = wb.createSheet("шим неж ИТО");     // имя — как в ТЗ
            Sheet resIto     = wb.createSheet("шум жил ИТО");
            Sheet autoDay    = wb.createSheet("шум авто день");
            Sheet autoNight  = wb.createSheet("шум авто ночь");
            Sheet site       = wb.createSheet("шум площадка");

            // Параметры страницы/поля и ширины колонок — одинаковые везде
            setupPage(day);       setupColumns(day);
            setupPage(night);     setupColumns(night);
            setupPage(nonResIto); setupColumns(nonResIto);
            setupPage(resIto);    setupColumns(resIto);
            setupPage(autoDay);   setupColumns(autoDay);
            setupPage(autoNight); setupColumns(autoNight);
            setupPage(site);      setupColumns(site);

            // Шапку отрисовываем ТОЛЬКО для «шум лифт день» (всё как вы задавали ранее)
            writeLiftHeader(wb, day, null);

            // Остальные листы — пока без наполнения (по вашему ТЗ на этом этапе)

            // Диалог сохранения — имя подставляем сразу
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
        ps.setLandscape(true);
        // поля указываются в дюймах
        sh.setMargin(Sheet.LeftMargin, 1.80 / 2.54);
        sh.setMargin(Sheet.RightMargin, 1.48 / 2.54);
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
        setCenter(sh, 2, 1, "№ точки измерения", verticalBorder);

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
        setText(r5, 4, "широкополосный", verticalBorder);
        setText(r5, 5, "тональный",       verticalBorder);
        setText(r5, 6, "постоянный",      verticalBorder);
        setText(r5, 7, "непостоянный",    verticalBorder);
        setText(r5, 8, "импульсный",      verticalBorder);

        // J3–R4
        merges.add(merge(sh, 2, 3, 9, 17));
        setCenter(sh, 2, 9,
                "Уровни звукового давления (дБ) ± U (дБ) в октавных полосах частот со среднегеометрическими частотами (Гц)",
                centerWrapBorder);

        // J5..R5
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

        // ===== A7–Y7: строка с датой/временем, выравнивание влево, без жирного =====
        if (dateLine != null && !dateLine.isBlank()) {
            setRowHeightCm(sh, 6, 0.53); // строка 7
            merge(sh, 6, 6, 0, 24);
            Row r7 = getOrCreateRow(sh, 6);
            Cell a7 = getOrCreateCell(r7, 0);
            a7.setCellValue(dateLine);

            CellStyle left8 = wb.createCellStyle();
            left8.cloneStyleFrom(leftNoWrap);
            left8.setFont(f8); // по умолчанию 8 пт
            a7.setCellStyle(left8);
        }
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
}
