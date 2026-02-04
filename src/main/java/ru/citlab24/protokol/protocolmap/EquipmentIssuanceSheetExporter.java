package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblBorders;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblCellMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGrid;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblWidth;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STBorder;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STFldCharType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJcTable;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblLayoutType;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STTblWidth;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class EquipmentIssuanceSheetExporter {
    private static final String ISSUANCE_SHEET_BASE_NAME = "лист выдачи приборов";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;
    private static final int TABLE_FONT_SIZE = 9;
    private static final int SIGNATURE_FONT_SIZE = 8;
    private static final String DATES_PREFIX = "2. Дата замеров:";
    private static final String PERFORMER_PREFIX = "3. Измерения провел, подпись:";
    private static final String OBJECT_PREFIX = "4. Наименование объекта:";
    private static final String INSTRUMENTS_PREFIX = "5.3. Приборы для измерения (используемое отметить):";
    private static final int NOISE_MERGED_DATE_LAST_COLUMN = 24;
    private static final int NOISE_PROTOCOL_MERGED_DATE_LAST_COLUMN = 23;
    private static final int NOISE_PROTOCOL_DATE_START_ROW = 5;
    private static final Pattern DATE_PATTERN = Pattern.compile("\\b\\d{2}\\.\\d{2}\\.(?:\\d{2}|\\d{4})\\b");

    private EquipmentIssuanceSheetExporter() {
    }

    static void generate(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        List<String> measurementDates = resolveMeasurementDates(mapFile);
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }
        String objectName = resolveObjectName(mapFile);
        String performer = resolveMeasurementPerformer(mapFile);
        List<InstrumentEntry> instruments = resolveInstruments(mapFile);

        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);
            File targetFile = resolveIssuanceSheetFile(mapFile, date, index, measurementDates.size());
            writeIssuanceSheet(targetFile, objectName, performer, instruments, date);
        }
    }

    public static List<File> resolveIssuanceSheetFiles(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }
        List<String> measurementDates = resolveMeasurementDates(mapFile);
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }
        List<File> files = new ArrayList<>();
        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);
            files.add(resolveIssuanceSheetFile(mapFile, date, index, measurementDates.size()));
        }
        return files;
    }

    private static void writeIssuanceSheet(File targetFile,
                                           String objectName,
                                           String performer,
                                           List<InstrumentEntry> instruments,
                                           String measurementDate) {
        try (XWPFDocument document = new XWPFDocument()) {
            applyStandardHeader(document);

            addCenteredParagraph(document, "Место использования прибора: " + safe(objectName), true);
            addCenteredParagraph(document, "Фамилия, инициалы лица, получившего приборы: ", true);
            addCenteredParagraph(document, safe(performer), false);
            addCenteredParagraph(document, "Лист выдачи приборов", true);

            int rows = 1 + instruments.size() + 2;
            XWPFTable table = document.createTable(rows, 5);
            configureIssuanceTableLayout(table);

            setTableCellText(table.getRow(0).getCell(0), "Наименование\nоборудования", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(1), "Зав. №", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(2),
                    "Получение прибора, отметка о проведении технического обслуживания и (или) проверки перед выдачей",
                    TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(3),
                    "Возврат оборудования, Отметка о проведении технического обслуживания и (или) проверки перед возвратом оборудования и после его транспортировки",
                    TABLE_FONT_SIZE);
            setTableCellText(table.getRow(0).getCell(4), "Примечание", TABLE_FONT_SIZE);

            int rowIndex = 1;
            for (InstrumentEntry instrument : instruments) {
                setTableCellText(table.getRow(rowIndex).getCell(0), safe(instrument.name), TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(1), safe(instrument.serialNumber), TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(2), "", TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(3), "", TABLE_FONT_SIZE);
                setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);
                rowIndex++;
            }

            setTableCellText(table.getRow(rowIndex).getCell(0), "Дата", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(1), "", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(2), safe(measurementDate), TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(3), "___.___.20____", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);
            rowIndex++;

            setTableCellText(table.getRow(rowIndex).getCell(0), "Подпись", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(1), "", TABLE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(2),
                    "__________\n(лица, получающего прибор)", SIGNATURE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(3),
                    "________\n(лица, ответственного за метрологическое обеспечение)",
                    SIGNATURE_FONT_SIZE);
            setTableCellText(table.getRow(rowIndex).getCell(4), "", TABLE_FONT_SIZE);

            addParagraphWithBreaks(document,
                    "Плановое обслуживание оборудования включает в себя осмотр на предмет отсутствия внешних повреждений, включение оборудования. " +
                            "Точный объем обслуживания определяется в эксплуатационной документации на оборудование, фиксируется в Плане обслуживания.\n" +
                            "При выдаче оборудования проводится проверка оборудования путем его осмотра на предмет целостности, отсутствия видимых дефектов, " +
                            "проверка уровня заряда аккумуляторных батарей (если применимо), наличия сигналов о неисправности при включении оборудования " +
                            "(если это не является составной частью технического обслуживания).\n" +
                            "Условные обозначения:+ - прибор получен (возвращен), техническое обслуживание и (или) промежуточная проверка (что предусмотрено) " +
                            "проведены, результат – прибор пригоден для дальнейшей работы.");

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    private static void addCenteredParagraph(XWPFDocument document, String text, boolean bold) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
        run.setBold(bold);
    }

    private static void addParagraphWithBreaks(XWPFDocument document, String text) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(TABLE_FONT_SIZE);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void setTableCellText(XWPFTableCell cell, String text, int fontSize) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);
        XWPFRun run = paragraph.createRun();
        run.setFontFamily(FONT_NAME);
        run.setFontSize(fontSize);

        String[] lines = text == null ? new String[]{""} : text.split("\\n", -1);
        for (int i = 0; i < lines.length; i++) {
            run.setText(lines[i]);
            if (i < lines.length - 1) {
                run.addBreak();
            }
        }
    }

    private static void applyStandardHeader(XWPFDocument document) {
        CTSectPr sectPr = document.getDocument().getBody().getSectPr();
        if (sectPr == null) {
            sectPr = document.getDocument().getBody().addNewSectPr();
        }

        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(document, sectPr);
        XWPFHeader header = policy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);

        XWPFTable table = header.createTable(3, 3);
        configureHeaderTableLikeTemplate(table);

        setHeaderCellText(table.getRow(0).getCell(0), "Испытательная лаборатория ООО «ЦИТ»");
        setHeaderCellText(table.getRow(0).getCell(1), "Лист выдачи приборов Ф39 ДП ИЛ 2-2023");
        setHeaderCellText(table.getRow(0).getCell(2), "Дата утверждения бланка формуляра: 01.01.2023г.");
        setHeaderCellText(table.getRow(1).getCell(2), "Редакция № 1");
        setHeaderCellPageCount(table.getRow(2).getCell(2));

        mergeCellsVertically(table, 0, 0, 2);
        mergeCellsVertically(table, 1, 0, 2);
    }

    private static void configureHeaderTableLikeTemplate(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(9639));

        CTJcTable jc = pr.isSetJc() ? pr.getJc() : pr.addNewJc();
        jc.setVal(STJcTable.CENTER);

        CTTblLayoutType layout = pr.isSetTblLayout() ? pr.getTblLayout() : pr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);

        CTTblBorders borders = pr.isSetTblBorders() ? pr.getTblBorders() : pr.addNewTblBorders();
        setBorder(borders.isSetTop() ? borders.getTop() : borders.addNewTop());
        setBorder(borders.isSetLeft() ? borders.getLeft() : borders.addNewLeft());
        setBorder(borders.isSetBottom() ? borders.getBottom() : borders.addNewBottom());
        setBorder(borders.isSetRight() ? borders.getRight() : borders.addNewRight());
        setBorder(borders.isSetInsideH() ? borders.getInsideH() : borders.addNewInsideH());
        setBorder(borders.isSetInsideV() ? borders.getInsideV() : borders.addNewInsideV());
        applyMinimalHeaderCellMargins(pr);

        CTTblGrid grid = ct.getTblGrid();
        if (grid == null) {
            grid = ct.addNewTblGrid();
        } else {
            while (grid.sizeOfGridColArray() > 0) {
                grid.removeGridCol(0);
            }
        }
        grid.addNewGridCol().setW(BigInteger.valueOf(2951));
        grid.addNewGridCol().setW(BigInteger.valueOf(3140));
        grid.addNewGridCol().setW(BigInteger.valueOf(3548));

        setCellWidth(table, 0, 0, 2951);
        setCellWidth(table, 1, 0, 2951);
        setCellWidth(table, 2, 0, 2951);

        setCellWidth(table, 0, 1, 3140);
        setCellWidth(table, 1, 1, 3140);
        setCellWidth(table, 2, 1, 3140);

        setCellWidth(table, 0, 2, 3548);
        setCellWidth(table, 1, 2, 3548);
        setCellWidth(table, 2, 2, 3548);
    }

    private static void configureIssuanceTableLayout(XWPFTable table) {
        CTTbl ct = table.getCTTbl();
        CTTblPr pr = ct.getTblPr() != null ? ct.getTblPr() : ct.addNewTblPr();

        CTTblWidth tblW = pr.isSetTblW() ? pr.getTblW() : pr.addNewTblW();
        tblW.setType(STTblWidth.DXA);
        tblW.setW(BigInteger.valueOf(9900));

        CTTblLayoutType layout = pr.isSetTblLayout() ? pr.getTblLayout() : pr.addNewTblLayout();
        layout.setType(STTblLayoutType.FIXED);

        CTTblGrid grid = ct.getTblGrid();
        if (grid == null) {
            grid = ct.addNewTblGrid();
        } else {
            while (grid.sizeOfGridColArray() > 0) {
                grid.removeGridCol(0);
            }
        }

        int[] widths = {1800, 1800, 2700, 2700, 900};
        for (int width : widths) {
            grid.addNewGridCol().setW(BigInteger.valueOf(width));
        }

        for (int rowIndex = 0; rowIndex < table.getNumberOfRows(); rowIndex++) {
            for (int colIndex = 0; colIndex < widths.length; colIndex++) {
                setCellWidth(table, rowIndex, colIndex, widths[colIndex]);
            }
        }
    }

    private static void setBorder(CTBorder border) {
        border.setVal(STBorder.SINGLE);
        border.setSz(BigInteger.valueOf(4));
        border.setSpace(BigInteger.ZERO);
        border.setColor("auto");
    }

    private static void applyMinimalHeaderCellMargins(CTTblPr pr) {
        CTTblCellMar cellMar = pr.isSetTblCellMar() ? pr.getTblCellMar() : pr.addNewTblCellMar();
        setCellMargin(cellMar.isSetTop() ? cellMar.getTop() : cellMar.addNewTop());
        setCellMargin(cellMar.isSetLeft() ? cellMar.getLeft() : cellMar.addNewLeft());
        setCellMargin(cellMar.isSetBottom() ? cellMar.getBottom() : cellMar.addNewBottom());
        setCellMargin(cellMar.isSetRight() ? cellMar.getRight() : cellMar.addNewRight());
    }

    private static void setCellMargin(CTTblWidth margin) {
        margin.setType(STTblWidth.DXA);
        margin.setW(BigInteger.ZERO);
    }

    private static void setCellWidth(XWPFTable table, int row, int col, int widthDxa) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        CTTcPr tcPr = cell.getCTTc().isSetTcPr() ? cell.getCTTc().getTcPr() : cell.getCTTc().addNewTcPr();
        CTTblWidth width = tcPr.isSetTcW() ? tcPr.getTcW() : tcPr.addNewTcW();
        width.setType(STTblWidth.DXA);
        width.setW(BigInteger.valueOf(widthDxa));
    }

    private static void setHeaderCellText(XWPFTableCell cell, String text) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);

        XWPFRun run = paragraph.createRun();
        run.setText(text != null ? text : "");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);
    }

    private static void setHeaderCellPageCount(XWPFTableCell cell) {
        cell.removeParagraph(0);
        XWPFParagraph paragraph = cell.addParagraph();
        paragraph.setAlignment(ParagraphAlignment.CENTER);
        paragraph.setSpacingAfter(0);
        paragraph.setSpacingBefore(0);

        XWPFRun run = paragraph.createRun();
        run.setText("Количество страниц: ");
        run.setFontFamily(FONT_NAME);
        run.setFontSize(FONT_SIZE);

        appendField(paragraph, "PAGE");
        XWPFRun separator = paragraph.createRun();
        separator.setText(" / ");
        separator.setFontFamily(FONT_NAME);
        separator.setFontSize(FONT_SIZE);

        appendField(paragraph, "NUMPAGES");
    }

    private static void appendField(XWPFParagraph paragraph, String instr) {
        XWPFRun runBegin = paragraph.createRun();
        runBegin.setFontFamily(FONT_NAME);
        runBegin.setFontSize(FONT_SIZE);
        runBegin.getCTR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);

        XWPFRun runInstr = paragraph.createRun();
        runInstr.setFontFamily(FONT_NAME);
        runInstr.setFontSize(FONT_SIZE);
        var ctText = runInstr.getCTR().addNewInstrText();
        ctText.setStringValue(instr);
        ctText.setSpace(SpaceAttribute.Space.PRESERVE);

        XWPFRun runSep = paragraph.createRun();
        runSep.setFontFamily(FONT_NAME);
        runSep.setFontSize(FONT_SIZE);
        runSep.getCTR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);

        XWPFRun runText = paragraph.createRun();
        runText.setFontFamily(FONT_NAME);
        runText.setFontSize(FONT_SIZE);
        runText.setText("1");

        XWPFRun runEnd = paragraph.createRun();
        runEnd.setFontFamily(FONT_NAME);
        runEnd.setFontSize(FONT_SIZE);
        runEnd.getCTR().addNewFldChar().setFldCharType(STFldCharType.END);
    }

    private static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (cell == null) {
                continue;
            }
            if (cell.getCTTc().getTcPr() == null) {
                cell.getCTTc().addNewTcPr();
            }
            if (rowIndex == fromRow) {
                cell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().getTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    private static List<String> resolveMeasurementDates(File mapFile) {
        List<String> noiseDates = resolveNoiseMeasurementDates(mapFile);
        if (!noiseDates.isEmpty()) {
            return noiseDates;
        }
        String rawDates = findValueByPrefix(mapFile, DATES_PREFIX);
        if (rawDates.isBlank()) {
            rawDates = findValueByPrefix(mapFile, "2. Дата замеров");
        }
        if (rawDates.isBlank()) {
            return Collections.emptyList();
        }
        String[] parts = rawDates.split(",");
        List<String> dates = new ArrayList<>();
        for (String part : parts) {
            String trimmed = part.trim();
            if (!trimmed.isEmpty()) {
                dates.add(trimmed);
            }
        }
        return dates;
    }

    private static List<String> resolveNoiseMeasurementDates(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() <= 1) {
                return Collections.emptyList();
            }
            DataFormatter formatter = new DataFormatter();
            var evaluator = workbook.getCreationHelper().createFormulaEvaluator();
            Set<String> dates = new LinkedHashSet<>();
            for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                if (sheet == null) {
                    continue;
                }
                for (org.apache.poi.ss.util.CellRangeAddress region : sheet.getMergedRegions()) {
                    if (!isNoiseDateRegion(region)) {
                        continue;
                    }
                    String text = readCellText(sheet, region.getFirstRow(), region.getFirstColumn(), formatter, evaluator);
                    if (!text.isEmpty()) {
                        addDatesFromText(text, dates);
                    }
                }
            }
            return new ArrayList<>(dates);
        } catch (Exception ignored) {
            return Collections.emptyList();
        }
    }

    private static boolean isNoiseDateRegion(org.apache.poi.ss.util.CellRangeAddress region) {
        return region.getFirstRow() == region.getLastRow()
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= NOISE_MERGED_DATE_LAST_COLUMN;
    }

    private static boolean isNoiseProtocolDateRegion(org.apache.poi.ss.util.CellRangeAddress region) {
        return region.getFirstRow() == region.getLastRow()
                && region.getFirstRow() >= NOISE_PROTOCOL_DATE_START_ROW
                && region.getFirstColumn() == 0
                && region.getLastColumn() >= NOISE_PROTOCOL_MERGED_DATE_LAST_COLUMN;
    }

    private static String readCellText(Sheet sheet,
                                       int rowIndex,
                                       int columnIndex,
                                       DataFormatter formatter,
                                       org.apache.poi.ss.usermodel.FormulaEvaluator evaluator) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return "";
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            return "";
        }
        return formatter.formatCellValue(cell, evaluator).trim();
    }

    private static void addDatesFromText(String text, Set<String> dates) {
        Matcher matcher = DATE_PATTERN.matcher(text);
        while (matcher.find()) {
            String date = matcher.group();
            if (!date.isEmpty()) {
                dates.add(date);
            }
        }
    }

    private static String resolveObjectName(File mapFile) {
        String value = findValueByPrefix(mapFile, OBJECT_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "4. Наименование объекта");
        }
        return value;
    }

    private static String resolveMeasurementPerformer(File mapFile) {
        String value = findValueByPrefix(mapFile, PERFORMER_PREFIX);
        if (value.isBlank()) {
            value = findValueByPrefix(mapFile, "3. Измерения провел, подпись");
        }
        return value;
    }

    private static List<InstrumentEntry> resolveInstruments(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return Collections.emptyList();
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            int startRow = findRowIndexWithText(sheet, formatter, INSTRUMENTS_PREFIX);
            if (startRow < 0) {
                return Collections.emptyList();
            }
            List<InstrumentEntry> instruments = new ArrayList<>();
            for (int rowIndex = startRow + 2; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    break;
                }
                if (rowContainsText(row, formatter, "6. Эскиз")) {
                    break;
                }
                String name = findFirstValueInRange(row, formatter, 0, 19);
                String serial = findFirstValueInRange(row, formatter, 20, 29);
                if (name.isEmpty() && serial.isEmpty()) {
                    if (isRowEmpty(row, formatter)) {
                        break;
                    }
                    continue;
                }
                instruments.add(new InstrumentEntry(name, serial));
            }
            return instruments;
        } catch (Exception ignored) {
            return Collections.emptyList();
        }
    }

    private static boolean rowContainsText(Row row, DataFormatter formatter, String needle) {
        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (text.contains(needle)) {
                return true;
            }
        }
        return false;
    }

    private static boolean isRowEmpty(Row row, DataFormatter formatter) {
        for (Cell cell : row) {
            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    private static int findRowIndexWithText(Sheet sheet, DataFormatter formatter, String needle) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.contains(needle)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static String findFirstValueInRange(Row row, DataFormatter formatter, int startCol, int endCol) {
        for (int col = startCol; col <= endCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                continue;
            }
            String text = formatter.formatCellValue(cell).trim();
            if (!text.isEmpty()) {
                return text;
            }
        }
        return "";
    }

    private static String findValueByPrefix(File mapFile, String prefix) {
        if (mapFile == null || !mapFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(mapFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            DataFormatter formatter = new DataFormatter();
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String text = formatter.formatCellValue(cell).trim();
                    if (text.startsWith(prefix)) {
                        String tail = text.substring(prefix.length()).trim();
                        if (!tail.isEmpty()) {
                            return trimLeadingPunctuation(tail);
                        }
                        Cell next = row.getCell(cell.getColumnIndex() + 1);
                        if (next != null) {
                            String nextText = formatter.formatCellValue(next).trim();
                            if (!nextText.isEmpty()) {
                                return trimLeadingPunctuation(nextText);
                            }
                        }
                    }
                }
            }
        } catch (Exception ignored) {
            return "";
        }
        return "";
    }

    private static String trimLeadingPunctuation(String value) {
        int index = 0;
        while (index < value.length() && !Character.isLetterOrDigit(value.charAt(index))) {
            index++;
        }
        return value.substring(index).trim();
    }

    private static File resolveIssuanceSheetFile(File mapFile, String date, int index, int total) {
        String name = ISSUANCE_SHEET_BASE_NAME;
        String safeDate = sanitizeFileComponent(date);
        if (total > 1) {
            if (safeDate.isBlank()) {
                name = name + " " + (index + 1);
            } else {
                name = name + " " + safeDate;
            }
        }
        return new File(mapFile.getParentFile(), name + ".docx");
    }

    private static String sanitizeFileComponent(String value) {
        if (value == null) {
            return "";
        }
        return value.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
    }

    private static String safe(String value) {
        return value == null ? "" : value;
    }

    private static final class InstrumentEntry {
        private final String name;
        private final String serialNumber;

        private InstrumentEntry(String name, String serialNumber) {
            this.name = name;
            this.serialNumber = serialNumber;
        }
    }
    static void generateForNoise(File sourceNoiseProtocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }

        // 1) Даты шумов берём из ИСХОДНОГО файла протокола шумов
        List<String> measurementDates = resolveNoiseMeasurementDatesFromProtocol(sourceNoiseProtocolFile);

        // 2) Если почему-то не нашли — оставляем один пустой лист (как было раньше)
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }

        // остальное (объект, исполнитель, приборы) берём из карты (там всё уже собрано)
        String objectName = resolveObjectName(mapFile);
        String performer = resolveMeasurementPerformer(mapFile);
        List<InstrumentEntry> instruments = resolveInstruments(mapFile);

        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);

            // ШУМЫ: дату в имя файла пишем всегда, если она есть
            File targetFile = resolveIssuanceSheetFileForNoise(mapFile, date, index, measurementDates.size());

            writeIssuanceSheet(targetFile, objectName, performer, instruments, date);
        }
    }

    static List<File> resolveIssuanceSheetFilesForNoise(File sourceNoiseProtocolFile, File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return Collections.emptyList();
        }

        List<String> measurementDates = resolveNoiseMeasurementDatesFromProtocol(sourceNoiseProtocolFile);
        if (measurementDates.isEmpty()) {
            measurementDates = List.of("");
        }

        List<File> files = new ArrayList<>();
        for (int index = 0; index < measurementDates.size(); index++) {
            String date = measurementDates.get(index);
            files.add(resolveIssuanceSheetFileForNoise(mapFile, date, index, measurementDates.size()));
        }
        return files;
    }

    private static List<String> resolveNoiseMeasurementDatesFromProtocol(File sourceNoiseProtocolFile) {
        if (sourceNoiseProtocolFile == null || !sourceNoiseProtocolFile.exists()) {
            return Collections.emptyList();
        }

        try (InputStream in = new FileInputStream(sourceNoiseProtocolFile);
             Workbook workbook = WorkbookFactory.create(in)) {

            // “смотрим на все вкладки кроме первой”
            if (workbook.getNumberOfSheets() <= 1) {
                return Collections.emptyList();
            }

            DataFormatter formatter = new DataFormatter();
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

            Set<String> dates = new LinkedHashSet<>();

            for (int sheetIndex = 1; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                Sheet sheet = workbook.getSheetAt(sheetIndex);
                if (sheet == null) {
                    continue;
                }

                // Ищем объединённую строку A..X (0..23) начиная с 6-й строки, читаем текст и выдёргиваем даты
                List<org.apache.poi.ss.util.CellRangeAddress> candidates = new ArrayList<>();
                for (org.apache.poi.ss.util.CellRangeAddress region : sheet.getMergedRegions()) {
                    if (isNoiseProtocolDateRegion(region)) {
                        candidates.add(region);
                    }
                }
                candidates.sort(java.util.Comparator.comparingInt(org.apache.poi.ss.util.CellRangeAddress::getFirstRow));

                for (org.apache.poi.ss.util.CellRangeAddress region : candidates) {
                    String text = readCellText(sheet,
                            region.getFirstRow(),
                            region.getFirstColumn(),
                            formatter,
                            evaluator);
                    if (!text.isEmpty()) {
                        addDatesFromText(text, dates);
                    }
                }
            }

            return new ArrayList<>(dates);

        } catch (Exception ignored) {
            return Collections.emptyList();
        }
    }

    private static File resolveIssuanceSheetFileForNoise(File mapFile, String date, int index, int total) {
        String name = ISSUANCE_SHEET_BASE_NAME;
        String safeDate = sanitizeFileComponent(date);

        // для шумов: если дата есть — всегда добавляем её к имени
        if (!safeDate.isBlank()) {
            name = name + " " + safeDate;
        } else if (total > 1) {
            // если даты нет, но листов несколько — нумеруем
            name = name + " " + (index + 1);
        }

        return new File(mapFile.getParentFile(), name + ".docx");
    }

}
