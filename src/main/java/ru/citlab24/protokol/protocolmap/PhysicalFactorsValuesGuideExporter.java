package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.util.CellReference;

import java.io.FileInputStream;
import java.util.Locale;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public final class PhysicalFactorsValuesGuideExporter {
    private static final String GUIDE_NAME = "Справка по значениям.docx";

    private PhysicalFactorsValuesGuideExporter() {
    }

    public static File generate(File mapFile) throws IOException {
        // старый контракт оставляем, но теперь это "читай и пиши из одного файла"
        return generate(mapFile, mapFile);
    }

    public static File generate(File sourceProtocolFile, File mapFile) throws IOException {
        File guideFile = resolveGuideFile(mapFile);
        if (guideFile == null) {
            return null;
        }

        try (XWPFDocument document = new XWPFDocument()) {
            RequestFormExporter.applyStandardHeader(document);

            addTitle(document, "Справка по значениям");
            addSectionTitle(document, "1. Вентиляция");
            addBullet(document, "Если значение площади сечения проема 0,008 — вставляем: Ø - 0,1 S= 0.008");
            addBullet(document, "Если значение площади сечения проема 0,012 — вставляем: Ø - 0,125 S= 0.012");
            addBullet(document, "Если значение площади сечения проема 0,016 — вставляем: Ø - 0,160 S= 0.016");
            addBullet(document, "Если значение площади сечения проема 0,020 — вставляем: Ø - 0,200 S= 0.020");
            addBullet(document, "Если значение площади сечения проема 0,010 — вставляем: 0,1х0,1 S= 0.01");
            addBullet(document, "Если значение площади сечения проема 0,060 — вставляем: 0,3х0,2 S= 0.06");
            addBullet(document, "Если значение площади сечения проема 0,080 — вставляем: 0,4х0,2 S= 0.08");
            addBullet(document, "Если значение площади сечения проема 0,100 — вставляем: 0,5х0,2 S= 0.1");
            addBullet(document, "Если значение площади сечения проема 0,150 — вставляем: 0,5х0,3 S= 0.15");

            addSectionTitle(document, "2. МЭД");
            addParagraph(document,
                    "На первом этапе ставьте любые значения из диапазона, указанного в карте, "
                            + "но обязательно должны присутствовать крайние точки (минимум и максимум). "
                            + "Например, при диапазоне 0,10–0,20 допускаются любые значения 0,10, 0,11, "
                            + "0,12, 0,13, 0,14, 0,15, 0,16, 0,17, 0,18, 0,19, 0,20, но 0,10 и 0,20 "
                            + "нужно указать обязательно. Всего требуется не менее 20 точек на этаж.");

            addSectionTitle(document, "3. Освещение на улице");
            addParagraph(document, "Таблица значений из протокола (лист «Иск освещение (2)», колонки K–AD):");
            addStreetLightingValuesTable(document, sourceProtocolFile);

            try (FileOutputStream out = new FileOutputStream(guideFile)) {
                document.write(out);
            }
        }

        return guideFile;
    }

    public static File resolveGuideFile(File mapFile) {
        if (mapFile == null || mapFile.getParentFile() == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), GUIDE_NAME);
    }

    private static void addTitle(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(14);
        run.setText(text);
    }

    private static void addSectionTitle(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setBold(true);
        run.setFontSize(12);
        run.setText(text);
    }

    private static void addBullet(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setFontSize(11);
        run.setText("• " + text);
    }

    private static void addParagraph(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun run = paragraph.createRun();
        run.setFontSize(11);
        run.setText(text);
    }

    private static void addStreetLightingValuesTable(XWPFDocument document, File sourceProtocolFile) throws IOException {
        List<List<String>> rows = readStreetLightingValuesRows(sourceProtocolFile);
        if (rows.isEmpty()) {
            addParagraph(document,
                    "Данные по листу «Иск освещение (2)» (K–AD) не найдены. " +
                            "Проверь: справка должна читать ИСХОДНЫЙ протокол, а не сформированную карту.");
            return;
        }

        final int columnCount = 20; // K..AD
        XWPFTable table = document.createTable(rows.size(), columnCount);

        for (int r = 0; r < rows.size(); r++) {
            List<String> values = rows.get(r);
            for (int c = 0; c < columnCount; c++) {
                String v = (c < values.size()) ? values.get(c) : "";
                setCellText(table.getRow(r).getCell(c), v);
            }
        }
    }

    private static void setCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        XWPFRun run = paragraph.createRun();
        run.setFontSize(11);
        run.setText(text);
    }

    private static void mergeCellsHorizontally(XWPFTable table, int row, int fromCol, int toCol) {
        for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(colIndex);
            CTTcPr cellProps = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
            if (colIndex == fromCol) {
                cellProps.addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cellProps.addNewHMerge().setVal(STMerge.CONTINUE);
                setCellText(cell, "");
            }
        }
    }
    private static List<List<String>> readStreetLightingValuesRows(File sourceProtocolFile) throws IOException {
        List<List<String>> rows = new ArrayList<>();
        if (sourceProtocolFile == null || !sourceProtocolFile.exists()) {
            return rows;
        }

        try (InputStream in = new FileInputStream(sourceProtocolFile);
             Workbook workbook = WorkbookFactory.create(in)) {

            Sheet sheet = workbook.getSheet("Иск освещение (2)");
            if (sheet == null) {
                return rows;
            }

            DataFormatter formatter = new DataFormatter(new Locale("ru", "RU"));
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            final int fromCol = CellReference.convertColStringToIndex("K");  // 10
            final int toCol   = CellReference.convertColStringToIndex("AD"); // 29

            int lastRow = sheet.getLastRowNum();

            // строго с 8-й строки Excel (0-based индекс 7)
            for (int rowIndex = 7; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                List<String> values = new ArrayList<>(toCol - fromCol + 1);
                boolean hasAny = false;

                for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
                    Cell cell = row.getCell(colIndex);
                    String v = formatCellValue(cell, formatter, evaluator);
                    values.add(v);
                    if (!v.isEmpty()) {
                        hasAny = true;
                    }
                }

                // по всем строкам: добавляем только те, где реально есть значения в K–AD
                if (hasAny) {
                    rows.add(values);
                }
            }
        }

        return rows;
    }

    private static List<StreetLightingRow> readStreetLightingRows(File mapFile) throws IOException {
        List<StreetLightingRow> rows = new ArrayList<>();
        if (mapFile == null || !mapFile.exists()) {
            return rows;
        }

        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {

            Sheet sheet = workbook.getSheet("Иск освещение (2)");
            if (sheet == null) {
                return rows;
            }

            // Важно: формулы в G считаем, а колонки берём по буквам (не руками индексами)
            DataFormatter formatter = new DataFormatter(new Locale("ru", "RU"));
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            final int placeCol = CellReference.convertColStringToIndex("C");    // C
            final int averageCol = CellReference.convertColStringToIndex("G");  // G
            final int fromCol = CellReference.convertColStringToIndex("K");     // K
            final int toCol = CellReference.convertColStringToIndex("AD");      // AD

            int lastRow = sheet.getLastRowNum();

            // Начинаем с 8-й строки Excel => индекс 7 (0-based)
            for (int rowIndex = 7; rowIndex <= lastRow; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break; // таблица закончилась
                }

                String place = formatCellValue(row.getCell(placeCol), formatter, evaluator);
                String average = formatCellValue(row.getCell(averageCol), formatter, evaluator);

                List<String> values = new ArrayList<>(toCol - fromCol + 1);
                boolean hasValues = false;

                for (int colIndex = fromCol; colIndex <= toCol; colIndex++) {
                    String value = formatCellValue(row.getCell(colIndex), formatter, evaluator);
                    values.add(value);
                    if (!value.isEmpty()) {
                        hasValues = true;
                    }
                }

                // По вашему требованию: "начиная с 8 пока есть значения"
                if (place.isEmpty() && average.isEmpty() && !hasValues) {
                    break;
                }

                rows.add(new StreetLightingRow(place, average, values));
            }
        }

        return rows;
    }
    private static String formatCellValue(Cell cell, DataFormatter formatter, FormulaEvaluator evaluator) {
        if (cell == null) {
            return "";
        }
        return formatter.formatCellValue(cell, evaluator).trim();
    }

    private record StreetLightingRow(String place, String average, List<String> values) {
    }
}
