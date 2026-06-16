package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import javax.imageio.ImageIO;
import java.awt.BasicStroke;
import java.awt.Color;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.Component;
import java.awt.Graphics2D;
import java.awt.RenderingHints;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.concurrent.ThreadLocalRandom;

final class AreaNoiseProtocolExporter {
    private static final String TEMPLATE = "/templates/area_noise_protocol_template.xlsx";
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String NOISE_METHOD_MI = "МИ Ш.13-2021 \"Методика измерений шума, инфразвука, воздушного\n"
            + "ультразвука на рабочих местах, в том числе рабочих местах\n"
            + "транспорта и объектов транспортной инфраструктуры, в\n"
            + "помещениях жилых, общественных и производственных\n"
            + "зданий, на селитебной и открытой территории.\" п. 12.3.3 / измерение шума";
    private static final String NOISE_METHOD_ECOFIZIKA = "Руководство по эксплуатации ПКДУ.411000.001.02 РЭ Шумомер-виброметр, "
            + "анализатор спектра Экофизика-110А п.7.1, п.7.2 (МИ ПКФ-12-006 методика измерений приложение "
            + "к руководству по эксплуатации п.2) / измерение шума";

    private AreaNoiseProtocolExporter() {
    }

    static void export(AreaProtocolPanel.AreaProtocolData data, Component parent) {
        try (InputStream input = AreaNoiseProtocolExporter.class.getResourceAsStream(TEMPLATE)) {
            if (input == null) {
                throw new IllegalStateException("Не найден шаблон протокола шума участка.");
            }
            try (Workbook workbook = WorkbookFactory.create(input)) {
                while (workbook.getNumberOfSheets() > 2) {
                    workbook.removeSheetAt(2);
                }
                Sheet titleSheet = workbook.getSheetAt(0);
                Sheet noiseSheet = workbook.getSheetAt(1);
                workbook.removePrintArea(workbook.getSheetIndex(titleSheet));
                fillTitleSheet(titleSheet, data);
                fillNoiseSheet(noiseSheet, data, formatLongDate(data.getProtocolDate()), buildProtocolNumber(data));
                setSuperscriptMinusSix(workbook);
                workbook.setActiveSheet(0);
                workbook.setSelectedTab(0);
                saveWorkbook(workbook, parent);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent,
                    "Не удалось экспортировать протокол ШУМ:\n" + ex.getMessage(),
                    "Экспорт протокола ШУМ",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private static void fillTitleSheet(Sheet sheet, AreaProtocolPanel.AreaProtocolData data) throws Exception {
        String approvalDate = formatLongDate(data.getProtocolDate());
        String protocolNumber = buildProtocolNumber(data);
        setCellText(sheet, "S7", approvalDate);
        setCellText(sheet, "A13", "Протокол испытаний № " + protocolNumber);
        updateFooter(sheet, approvalDate, protocolNumber);
        setCellText(sheet, "B16", "Наименование и контактные данные заявителя (заказчика): "
                + nullToEmpty(data.getCustomerNameAndContacts()));
        setCellText(sheet, "B18", "Юридический адрес заказчика: " + nullToEmpty(data.getCustomerLegalAddress()));
        setCellText(sheet, "B20", "Фактический адрес заказчика: " + nullToEmpty(data.getCustomerActualAddress()));
        setCellText(sheet, "B22", "Наименование предприятия, организации, объекта, где производились измерения: "
                + nullToEmpty(data.getObjectName()));
        setCellText(sheet, "B24", "Адрес предприятия (объекта): " + buildObjectAddressText(data));
        setCellText(sheet, "B26", buildReasonText(data));
        setCellText(sheet, "B28", "Измерения проводились в присутствии представителя заказчика: "
                + nullToEmpty(data.getRepresentative()));
        if (nullToEmpty(data.getRepresentative()).isBlank()) {
            highlightRange(sheet, 27, 27, 1, 25);
        }
        setCellText(sheet, "B30", "Показатели, по которым проводились измерения: эквивалентный уровень звука, максимальный уровень звука.");
        setCellText(sheet, "B32", "Регистрационный номер карты замеров: " + buildMeasurementCardNumber(data));
        setCellText(sheet, "N44", "± 0,13 кПа\n(±1 мм рт. ст.)");
        setCellText(sheet, "P54", resolveNoiseMethod(data));
        setCellText(sheet, "B56", buildAdditionalInfoText(data));
        insertNoiseSketch(sheet, data.getNoiseSketchImage());
        updateSketchLegend(sheet);
    }

    private static void insertNoiseSketch(Sheet sheet, BufferedImage image) throws IOException {
        if (image == null) {
            return;
        }
        int sketchTitleRow = findRowContainingText(sheet, "Эскиз (ситуационный план)");
        if (sketchTitleRow < 0) {
            return;
        }
        int firstRow = sketchTitleRow + 2;
        int lastRow = firstRow + 14;
        int firstCol = 1;
        int lastCol = 12;
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            if (sheet.getRow(rowIndex) == null) {
                sheet.createRow(rowIndex);
            }
        }
        CellRangeAddress region = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        removeOverlappingMergedRegions(sheet, region);
        sheet.addMergedRegion(region);

        byte[] imageBytes = toPngBytes(image);
        Workbook workbook = sheet.getWorkbook();
        int pictureIndex = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        CreationHelper helper = workbook.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(firstCol);
        anchor.setRow1(firstRow);
        Picture picture = drawing.createPicture(anchor, pictureIndex);
        double widthScale = regionWidthPixels(sheet, firstCol, lastCol) / image.getWidth();
        double heightScale = regionHeightPixels(sheet, firstRow, lastRow) / image.getHeight();
        double scale = Math.min(widthScale, heightScale);
        if (scale > 0d) {
            picture.resize(scale);
        }
    }

    private static byte[] toPngBytes(BufferedImage image) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ImageIO.write(image, "png", output);
        return output.toByteArray();
    }

    private static void updateSketchLegend(Sheet sheet) throws IOException {
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        addLegendIcon(sheet, drawing, "O64", createBoundaryLegendIcon());
    }

    private static void addLegendIcon(Sheet sheet, Drawing<?> drawing, String cellAddress, BufferedImage icon)
            throws IOException {
        Workbook workbook = sheet.getWorkbook();
        int pictureIndex = workbook.addPicture(toPngBytes(icon), Workbook.PICTURE_TYPE_PNG);
        CellReference ref = new CellReference(cellAddress);
        ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
        anchor.setCol1(ref.getCol());
        anchor.setRow1(ref.getRow());
        Picture picture = drawing.createPicture(anchor, pictureIndex);
        picture.resize(1d);
    }

    private static BufferedImage createBoundaryLegendIcon() {
        BufferedImage image = new BufferedImage(72, 22, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2 = image.createGraphics();
        try {
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(
                    3f,
                    BasicStroke.CAP_BUTT,
                    BasicStroke.JOIN_MITER,
                    10f,
                    new float[]{12f, 9f},
                    0f
            ));
            g2.drawLine(8, 11, 68, 11);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private static void fillNoiseSheet(Sheet sheet, AreaProtocolPanel.AreaProtocolData data,
                                       String approvalDate, String protocolNumber) {
        int pointCount = Math.max(1, data.getNoisePointCount());
        updateFooter(sheet, approvalDate, protocolNumber);
        removeRowBreaks(sheet);
        adjustNoiseRows(sheet, pointCount);
        setCellText(sheet, "A7", "Дата, время проведения измерений " + findNoiseDate(data)
                + " c ____ до ____");
        for (int index = 0; index < pointCount; index++) {
            int rowIndex = 7 + index;
            Row row = getOrCreateRow(sheet, rowIndex);
            row.setZeroHeight(false);
            setCell(row, 0, index + 1);
            setCell(row, 1, "т" + (index + 1));
            setCell(row, 4, "+");
            setCell(row, 5, "-");
            setCell(row, 6, "-");
            setCell(row, 7, "+");
            setCell(row, 8, "-");
            for (int col = 9; col <= 18; col++) {
                setCell(row, col, "-");
            }
            setCell(row, 19, randomNoiseValue(data.getNoiseEquivalentMinValue(), data.getNoiseEquivalentMaxValue()));
            setCell(row, 22, randomNoiseValue(data.getNoiseMaxLevelMinValue(), data.getNoiseMaxLevelMaxValue()));
        }
        int visibleDataRows = Math.min(pointCount, 8);
        for (int index = visibleDataRows; index < 8; index++) {
            Row row = sheet.getRow(7 + index);
            if (row != null) {
                clearRow(row, 0, 22);
                row.setZeroHeight(true);
            }
        }
        rebuildNoiseDataMerges(sheet, pointCount);
    }

    private static void adjustNoiseRows(Sheet sheet, int pointCount) {
        int existingRows = 8;
        int extraRows = pointCount - existingRows;
        if (extraRows > 0) {
            sheet.shiftRows(15, sheet.getLastRowNum(), extraRows, true, false);
            Row templateRow = sheet.getRow(14);
            for (int rowIndex = 15; rowIndex < 15 + extraRows; rowIndex++) {
                copyRowStructure(sheet, templateRow, rowIndex);
            }
        }
    }

    private static void copyRowStructure(Sheet sheet, Row templateRow, int rowIndex) {
        Row row = getOrCreateRow(sheet, rowIndex);
        if (templateRow != null) {
            row.setHeight(templateRow.getHeight());
        }
        for (int col = 0; col <= 22; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            Cell templateCell = templateRow == null ? null : templateRow.getCell(col);
            if (templateCell != null) {
                cell.setCellStyle(templateCell.getCellStyle());
            }
            cell.setCellValue("");
        }
    }

    private static void rebuildNoiseDataMerges(Sheet sheet, int pointCount) {
        int lastDataRow = 7 + Math.max(pointCount, 8) - 1;
        removeDataMerges(sheet, 7, lastDataRow);
        for (int rowIndex = 7; rowIndex < 7 + pointCount; rowIndex++) {
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 19, 21));
        }
        if (pointCount > 1) {
            sheet.addMergedRegion(new CellRangeAddress(7, 7 + pointCount - 1, 2, 2));
            sheet.addMergedRegion(new CellRangeAddress(7, 7 + pointCount - 1, 3, 3));
        }
        setCell(sheet.getRow(7), 2, "Земельный участок");
        setCell(sheet.getRow(7), 3, "Суммарные источники шума (в т.ч. движение авто- и железнодорожного транспорта)");
    }

    private static void removeDataMerges(Sheet sheet, int firstRow, int lastRow) {
        for (int index = sheet.getNumMergedRegions() - 1; index >= 0; index--) {
            CellRangeAddress region = sheet.getMergedRegion(index);
            if (region.getFirstRow() < firstRow || region.getLastRow() > lastRow) {
                continue;
            }
            boolean sourceColumns = region.getFirstColumn() >= 2 && region.getLastColumn() <= 3;
            boolean resultColumns = region.getFirstColumn() >= 19 && region.getLastColumn() <= 21;
            if (sourceColumns || resultColumns) {
                sheet.removeMergedRegion(index);
            }
        }
    }

    private static void removeRowBreaks(Sheet sheet) {
        for (int rowBreak : sheet.getRowBreaks()) {
            sheet.removeRowBreak(rowBreak);
        }
    }

    private static void setSuperscriptMinusSix(Workbook workbook) {
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() != CellType.STRING) {
                        continue;
                    }
                    String text = cell.getStringCellValue();
                    if (!text.contains("10-6")) {
                        continue;
                    }
                    RichTextString richText = workbook.getCreationHelper().createRichTextString(text);
                    Font baseFont = workbook.createFont();
                    baseFont.setFontName("Arial");
                    baseFont.setFontHeightInPoints((short) 8);
                    richText.applyFont(0, text.length(), baseFont);
                    Font superFont = workbook.createFont();
                    superFont.setFontName("Arial");
                    superFont.setFontHeightInPoints((short) 8);
                    superFont.setTypeOffset(Font.SS_SUPER);
                    int index = text.indexOf("-6");
                    while (index >= 0) {
                        richText.applyFont(index, index + 2, superFont);
                        index = text.indexOf("-6", index + 2);
                    }
                    cell.setCellValue(richText);
                }
            }
        }
    }

    private static String buildProtocolNumber(AreaProtocolPanel.AreaProtocolData data) {
        return safeApplicationNumber(data) + "-2-Ф/" + twoDigitYear(data);
    }

    private static void updateFooter(Sheet sheet, String approvalDate, String protocolNumber) {
        sheet.getFooter().setRight("&\"Arial,обычный\"&9&K000000Протокол от "
                + nullToEmpty(approvalDate)
                + " № " + nullToEmpty(protocolNumber)
                + "\nОбщее количество страниц &N Страница &P");
    }

    private static String buildMeasurementCardNumber(AreaProtocolPanel.AreaProtocolData data) {
        return safeApplicationNumber(data) + "-2К";
    }

    private static String resolveNoiseMethod(AreaProtocolPanel.AreaProtocolData data) {
        String method = nullToEmpty(data.getNoiseMethod());
        if (method.contains("Экофизика")) {
            return NOISE_METHOD_ECOFIZIKA;
        }
        if (method.contains("МИ Ш.13")) {
            return NOISE_METHOD_MI;
        }
        return method.isBlank() ? NOISE_METHOD_MI : method;
    }

    private static String safeApplicationNumber(AreaProtocolPanel.AreaProtocolData data) {
        String number = nullToEmpty(data.getApplicationNumber());
        return number.isBlank() ? "____" : number;
    }

    private static String twoDigitYear(AreaProtocolPanel.AreaProtocolData data) {
        LocalDate date = parseDate(data.getProtocolDate());
        if (date == null) {
            date = parseDate(data.getApplicationDate());
        }
        if (date == null) {
            date = LocalDate.now();
        }
        return String.format(Locale.ROOT, "%02d", date.getYear() % 100);
    }

    private static String buildObjectAddressText(AreaProtocolPanel.AreaProtocolData data) {
        String address = nullToEmpty(data.getObjectAddress());
        String area = nullToEmpty(data.getAreaText());
        if (area.isBlank()) {
            return address;
        }
        if (address.toLowerCase(new Locale("ru", "RU")).contains("площад")) {
            return address;
        }
        return address + " площадью " + area + " кв. м.";
    }

    private static String buildReasonText(AreaProtocolPanel.AreaProtocolData data) {
        return "Основание для измерений: договор № " + nullToEmpty(data.getContractNumber())
                + " от " + nullToEmpty(data.getContractDate())
                + ", заявка № " + nullToEmpty(data.getApplicationNumber())
                + " от " + nullToEmpty(data.getApplicationDate());
    }

    private static String buildAdditionalInfoText(AreaProtocolPanel.AreaProtocolData data) {
        return "Дополнительные сведения (характеристика объекта): " + buildNoiseWeatherText(data)
                + " Эскиз предоставлен заказчиком. При измерениях на объекте заказчик самостоятельно определяет "
                + "расположение точек в соответствии с эскизом.";
    }

    private static String buildNoiseWeatherText(AreaProtocolPanel.AreaProtocolData data) {
        List<String> parts = new ArrayList<>();
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (row == null || !row.isNoiseSelected()) {
                continue;
            }
            String temperature = buildTemperatureText(row);
            if (temperature.isBlank()) {
                temperature = "от ____ °С до ____ °С";
            }
            parts.add("Температура воздуха " + temperature
                    + ", относительная влажность от ____ % до ____ %, атмосферное давление ____ мм рт. ст., "
                    + "скорость движения ветра от ____ м/с до ____ м/с, без осадков.");
        }
        if (parts.isEmpty()) {
            return "Температура воздуха от ____ °С до ____ °С, относительная влажность от ____ % до ____ %, "
                    + "атмосферное давление ____ мм рт. ст., скорость движения ветра от ____ м/с до ____ м/с, без осадков.";
        }
        return String.join(" ", parts);
    }

    private static String buildTemperatureText(AreaProtocolPanel.MeasurementRowData row) {
        String start = nullToEmpty(row.getTempOutsideStart());
        String end = nullToEmpty(row.getTempOutsideEnd());
        if (!start.isBlank() && !end.isBlank()) {
            return "от " + start + " °С до " + end + " °С";
        }
        if (!start.isBlank()) {
            return start + " °С";
        }
        if (!end.isBlank()) {
            return end + " °С";
        }
        return "";
    }

    private static String findNoiseDate(AreaProtocolPanel.AreaProtocolData data) {
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (row != null && row.isNoiseSelected() && !nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (row != null && !nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        return nullToEmpty(data.getProtocolDate());
    }

    private static String randomNoiseValue(String minText, String maxText) {
        double min = parseDecimal(minText);
        double max = parseDecimal(maxText);
        if (min <= 0d || max <= 0d) {
            return "____ ± 2,3";
        }
        if (min > max) {
            double temp = min;
            min = max;
            max = temp;
        }
        double value = max <= min ? min : ThreadLocalRandom.current().nextDouble(min, max);
        return formatOneDecimal(value) + " ± 2,3";
    }

    private static double parseDecimal(String value) {
        if (value == null || value.isBlank()) {
            return 0d;
        }
        String normalized = value.trim().replace(" ", "").replace(",", ".");
        try {
            return Double.parseDouble(normalized);
        } catch (NumberFormatException ex) {
            return 0d;
        }
    }

    private static String formatOneDecimal(double value) {
        return String.format(Locale.ROOT, "%.1f", value).replace(".", ",");
    }

    private static int findRowContainingText(Sheet sheet, String text) {
        DataFormatter formatter = new DataFormatter(new Locale("ru", "RU"));
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (formatter.formatCellValue(cell).contains(text)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static void removeOverlappingMergedRegions(Sheet sheet, CellRangeAddress target) {
        for (int index = sheet.getNumMergedRegions() - 1; index >= 0; index--) {
            CellRangeAddress existing = sheet.getMergedRegion(index);
            if (regionsIntersect(existing, target)) {
                sheet.removeMergedRegion(index);
            }
        }
    }

    private static boolean regionsIntersect(CellRangeAddress first, CellRangeAddress second) {
        return first.getFirstRow() <= second.getLastRow()
                && first.getLastRow() >= second.getFirstRow()
                && first.getFirstColumn() <= second.getLastColumn()
                && first.getLastColumn() >= second.getFirstColumn();
    }

    private static double regionWidthPixels(Sheet sheet, int firstCol, int lastCol) {
        double width = 0d;
        for (int col = firstCol; col <= lastCol; col++) {
            width += sheet.getColumnWidthInPixels(col);
        }
        return width;
    }

    private static double regionHeightPixels(Sheet sheet, int firstRow, int lastRow) {
        double height = 0d;
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            float heightPoints = row == null ? sheet.getDefaultRowHeightInPoints() : row.getHeightInPoints();
            height += heightPoints * 96d / 72d;
        }
        return height;
    }

    private static Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? sheet.createRow(rowIndex) : row;
    }

    private static void setCellText(Sheet sheet, String address, String value) {
        getCell(sheet, address).setCellValue(value == null ? "" : value);
    }

    private static void highlightRange(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = getOrCreateRow(sheet, rowIndex);
            for (int colIndex = firstCol; colIndex <= lastCol; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }
                cell.setCellStyle(highlightStyle(sheet.getWorkbook(), cell.getCellStyle()));
            }
        }
    }

    private static CellStyle highlightStyle(Workbook workbook, CellStyle baseStyle) {
        CellStyle style = workbook.createCellStyle();
        if (baseStyle != null) {
            style.cloneStyleFrom(baseStyle);
        }
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private static void setCell(Row row, int column, String value) {
        Cell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        cell.setCellValue(value == null ? "" : value);
    }

    private static void setCell(Row row, int column, int value) {
        Cell cell = row.getCell(column);
        if (cell == null) {
            cell = row.createCell(column);
        }
        cell.setCellValue(value);
    }

    private static Cell getCell(Sheet sheet, String address) {
        CellReference ref = new CellReference(address);
        Row row = getOrCreateRow(sheet, ref.getRow());
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) {
            cell = row.createCell(ref.getCol());
        }
        return cell;
    }

    private static void clearRow(Row row, int firstCol, int lastCol) {
        for (int col = firstCol; col <= lastCol; col++) {
            Cell cell = row.getCell(col);
            if (cell != null) {
                cell.setCellValue("");
            }
        }
    }

    private static String formatLongDate(String value) {
        LocalDate date = parseDate(value);
        if (date == null) {
            return "";
        }
        String[] months = {
                "января", "февраля", "марта", "апреля", "мая", "июня",
                "июля", "августа", "сентября", "октября", "ноября", "декабря"
        };
        return date.getDayOfMonth() + " " + months[date.getMonthValue() - 1] + " " + date.getYear() + " г.";
    }

    private static LocalDate parseDate(String value) {
        if (value == null || value.isBlank()) {
            return null;
        }
        try {
            return LocalDate.parse(value.trim(), DATE_FORMAT);
        } catch (Exception ex) {
            return null;
        }
    }

    private static String nullToEmpty(String value) {
        return value == null ? "" : value.trim();
    }

    private static void saveWorkbook(Workbook workbook, Component parent) throws Exception {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Сохранить протокол ШУМ участка");
        chooser.setSelectedFile(new File("Протокол_участок_шум.xlsx"));
        chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        if (chooser.showSaveDialog(parent) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File file = chooser.getSelectedFile();
        if (!file.getName().toLowerCase(Locale.ROOT).endsWith(".xlsx")) {
            file = new File(file.getAbsolutePath() + ".xlsx");
        }
        try (FileOutputStream out = new FileOutputStream(file)) {
            workbook.write(out);
        }
        JOptionPane.showMessageDialog(parent,
                "Файл сохранён:\n" + file.getAbsolutePath(),
                "Экспорт завершён",
                JOptionPane.INFORMATION_MESSAGE);
    }
}
