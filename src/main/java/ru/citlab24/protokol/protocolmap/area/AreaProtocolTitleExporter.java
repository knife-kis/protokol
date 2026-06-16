package ru.citlab24.protokol.protocolmap.area;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;

import javax.imageio.ImageIO;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.BasicStroke;
import java.awt.Color;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.awt.Component;
import java.awt.Graphics2D;
import java.awt.Polygon;
import java.awt.RenderingHints;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.concurrent.ThreadLocalRandom;

final class AreaProtocolTitleExporter {
    private static final String TEMPLATE = "/templates/area_protocol_title.xlsx";
    private static final String TABS_TEMPLATE = "/templates/area_protocol_rad_tabs_template.xlsx";
    private static final DateTimeFormatter DATE_FORMAT = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    private static final String PPR_INDICATOR =
            "плотность потока радона (ППР) с поверхности грунта, среднее значение плотности потока радона с поверхности почвы";
    private static final String PPR_ND_INDICATOR =
            "Плотность потока радона (ППР) с поверхности, среднее значение плотности потока радона с поверхности почвы";
    private static final String GAMMA_INDICATOR = "мощность дозы гамма-излучения";
    private static final String GAMMA_ND_INDICATOR = "Мощность дозы гамма-излучения";
    private static final String SANPIN =
            "СанПиН 2.6.4115-25\n" +
                    "\"Санитарно-эпидемиологические требования в области радиационной безопасности населения при обращении источников ионизирующего излучения\"";
    private static final String METHOD =
            "МР 2.6.1.0361-24 п.IV \"Радиационный контроль земельных участков, предназначенных под строительство жилых домов, зданий и сооружений общественного и производственного назначения, а также прилегающей к зданиям и сооружениям территории и территории общего пользования\" / Прочие методы радиационных исследований (испытаний)";

    private AreaProtocolTitleExporter() {
    }

    static void export(AreaProtocolPanel.AreaProtocolData data, Component parent) {
        try (InputStream input = AreaProtocolTitleExporter.class.getResourceAsStream(TEMPLATE)) {
            if (input == null) {
                throw new IllegalStateException("Не найден шаблон титульного листа участка.");
            }
            try (Workbook workbook = WorkbookFactory.create(input)) {
                keepOnlyFirstSheet(workbook);
                Sheet sheet = workbook.getSheetAt(0);
                workbook.setSheetName(0, "Титульный");
                fillSheet(sheet, data);
                addMeasurementSheets(workbook, data);
                saveWorkbook(workbook, parent);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(parent,
                    "Ошибка экспорта протокола участка: " + ex.getMessage(),
                    "Экспорт протокола участка",
                    JOptionPane.ERROR_MESSAGE);
        }
    }

    private static void fillSheet(Sheet sheet, AreaProtocolPanel.AreaProtocolData data) throws IOException {
        setupPrint(sheet);
        String approvalDate = formatLongDate(data.getProtocolDate());
        String protocolNumber = buildProtocolNumber(data);
        setCellText(sheet, "S7", approvalDate);
        setCellText(sheet, "A14", "Протокол испытаний № " + protocolNumber);
        updateFooter(sheet, approvalDate, protocolNumber);
        setCellText(sheet, "B17", "Наименование и контактные данные заявителя (заказчика): "
                + nullToEmpty(data.getCustomerNameAndContacts()));
        setCellText(sheet, "B19", "Юридический адрес заказчика: " + nullToEmpty(data.getCustomerLegalAddress()));
        setCellText(sheet, "B21", "Фактический адрес заказчика: " + nullToEmpty(data.getCustomerActualAddress()));
        setCellText(sheet, "B23", "Наименование предприятия, организации, объекта, где производились измерения: "
                + nullToEmpty(data.getObjectName()));
        setCellText(sheet, "B25", "Адрес предприятия (объекта): " + nullToEmpty(data.getObjectAddress()));
        setCellText(sheet, "B27", buildReasonText(data));
        setCellText(sheet, "B29", "Измерения проводились в присутствии представителя заказчика: "
                + nullToEmpty(data.getRepresentative()));
        setCellText(sheet, "B31", "Показатели, по которым проводились измерения: " + buildIndicatorsText(data));
        setCellText(sheet, "B33", "Регистрационный номер карты замеров и(или) журнал регистрации результатов определения плотности потока радона: 1/02.02.26-1К и 1/2026");
        highlightRange(sheet, 32, 32, 1, 25);

        setSuperscriptMinusSix(sheet);
        setCellText(sheet, "N44", "± 0,13 кПа\n(±1 мм рт. ст.)");
        setCellText(sheet, "A51", "Измерение напряжения и частоты переменного тока");

        setCellText(sheet, "A57", buildNdIndicatorText(data));
        setCellText(sheet, "F57", SANPIN);
        setCellText(sheet, "P57", METHOD);
        setCellText(sheet, "A58", "");
        setCellText(sheet, "F58", "");
        setCellText(sheet, "P58", "");
        hideSecondNormativeRow(sheet);

        setCellText(sheet, "B61", buildAdditionalInfoText(data));
        highlightRange(sheet, 60, 62, 1, 25);
        setCellText(sheet, "B64", data.isMedSelected() ? buildGammaSurveyText(data) : "");
        sheet.setRowBreak(67);
        insertRadiationSketch(sheet, data.getRadiationSketchImage());
        updateSketchLegend(sheet);
    }

    private static void insertRadiationSketch(Sheet sheet, BufferedImage image) throws IOException {
        if (image == null) {
            return;
        }
        int sketchTitleRow = findRowContainingText(sheet, "Эскиз (ситуационный план)");
        if (sketchTitleRow < 0) {
            return;
        }
        int firstRow = sketchTitleRow + 2;
        int lastRow = firstRow + 29;
        int firstCol = 1;
        int lastCol = 17;
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
        setCellText(sheet, "U83", "Профили пешеходжной гамма-съемки (на высоте до 0,3 м)");
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        addLegendIcon(sheet, drawing, "S74", createBoundaryLegendIcon());
        addLegendIcon(sheet, drawing, "S78", createGammaControlLegendIcon());
        addLegendIcon(sheet, drawing, "S83", createGammaProfileLegendIcon());
        addLegendIcon(sheet, drawing, "S87", createPprLegendIcon());
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
        BufferedImage image = createLegendIconCanvas(72, 22);
        Graphics2D g2 = image.createGraphics();
        try {
            setupLegendGraphics(g2);
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

    private static BufferedImage createGammaControlLegendIcon() {
        BufferedImage image = createLegendIconCanvas(62, 52);
        Graphics2D g2 = image.createGraphics();
        try {
            setupLegendGraphics(g2);
            int x = 11;
            int y = 4;
            int width = 38;
            int height = 44;
            g2.setColor(Color.WHITE);
            g2.fillRect(x, y, width, height);
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(2f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));
            g2.drawRect(x, y, width, height);
            drawCenteredLegendNumber(g2, "1", x, y, width, height);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private static BufferedImage createPprLegendIcon() {
        BufferedImage image = createLegendIconCanvas(66, 56);
        Graphics2D g2 = image.createGraphics();
        try {
            setupLegendGraphics(g2);
            Polygon diamond = new Polygon(
                    new int[]{33, 60, 33, 6},
                    new int[]{4, 28, 52, 28},
                    4
            );
            g2.setColor(Color.WHITE);
            g2.fillPolygon(diamond);
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(2f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));
            g2.drawPolygon(diamond);
            drawCenteredLegendNumber(g2, "1", 6, 4, 54, 48);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private static BufferedImage createGammaProfileLegendIcon() {
        BufferedImage image = createLegendIconCanvas(68, 42);
        Graphics2D g2 = image.createGraphics();
        try {
            setupLegendGraphics(g2);
            int circleX = 3;
            int circleY = 7;
            int circleSize = 29;
            int centerY = circleY + circleSize / 2;
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(1.8f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER));
            g2.drawLine(circleX + circleSize, centerY, 64, centerY);
            g2.setColor(Color.WHITE);
            g2.fillOval(circleX, circleY, circleSize, circleSize);
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(1.4f));
            g2.drawOval(circleX, circleY, circleSize, circleSize);
            drawCenteredLegendNumber(g2, "1", circleX, circleY, circleSize, circleSize);
        } finally {
            g2.dispose();
        }
        return image;
    }

    private static BufferedImage createLegendIconCanvas(int width, int height) {
        return new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
    }

    private static void setupLegendGraphics(Graphics2D g2) {
        g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
    }

    private static void drawCenteredLegendNumber(Graphics2D g2, String text, int x, int y, int width, int height) {
        g2.setColor(Color.BLACK);
        g2.setFont(new java.awt.Font("Arial", java.awt.Font.PLAIN, 18));
        java.awt.FontMetrics metrics = g2.getFontMetrics();
        int textX = x + (width - metrics.stringWidth(text)) / 2;
        int textY = y + (height + metrics.getAscent() - metrics.getDescent()) / 2;
        g2.drawString(text, textX, textY);
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

    private static void setupPrint(Sheet sheet) {
        sheet.getWorkbook().removePrintArea(sheet.getWorkbook().getSheetIndex(sheet));
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setPaperSize(PrintSetup.A4_PAPERSIZE);
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
    }

    private static void updateFooter(Sheet sheet, String approvalDate, String protocolNumber) {
        sheet.getFooter().setRight("&\"Arial,обычный\"&9&K000000Протокол от "
                + nullToEmpty(approvalDate)
                + " № " + nullToEmpty(protocolNumber)
                + "\nОбщее количество страниц &N Страница &P");
    }

    private static void hideSecondNormativeRow(Sheet sheet) {
        Row row = sheet.getRow(57);
        if (row != null) {
            row.setZeroHeight(true);
        }
    }

    private static void keepOnlyFirstSheet(Workbook workbook) {
        while (workbook.getNumberOfSheets() > 1) {
            workbook.removeSheetAt(1);
        }
        workbook.setActiveSheet(0);
        workbook.setSelectedTab(0);
    }

    private static void addMeasurementSheets(Workbook workbook, AreaProtocolPanel.AreaProtocolData data)
            throws IOException {
        try (InputStream input = AreaProtocolTitleExporter.class.getResourceAsStream(TABS_TEMPLATE)) {
            if (input == null) {
                throw new IllegalStateException("Не найден шаблон вкладок МЭД/ППР.");
            }
            try (Workbook sourceWorkbook = WorkbookFactory.create(input)) {
                Sheet sourceMedSheet = sourceWorkbook.getSheet("МЭД");
                Sheet sourcePprSheet = sourceWorkbook.getSheet("ППР");
                if (data.isMedSelected()) {
                    Sheet medSheet = copySheetFragment(sourceMedSheet, workbook, "МЭД", 0, 5, 0, 12);
                    setupMeasurementSheet(medSheet);
                    setupMedColumnWidths(medSheet);
                    int nextRow = fillMedRows(medSheet, sourceMedSheet, data);
                    if (!data.isPprSelected()) {
                        appendAreaProtocolFooter(medSheet, sourcePprSheet, nextRow, data);
                    }
                    superscriptSquareMeters(medSheet);
                }
                if (data.isPprSelected()) {
                    Sheet pprSheet = copySheetFragment(sourcePprSheet, workbook, "ППР", 0, 3, 0, 13);
                    setupMeasurementSheet(pprSheet);
                    setupPprSheet(pprSheet);
                    int nextRow = fillPprRows(pprSheet, sourcePprSheet, data);
                    appendAreaProtocolFooter(pprSheet, sourcePprSheet, nextRow, data);
                    superscriptSquareMeters(pprSheet);
                }
            }
        }
    }

    private static Sheet copySheetFragment(Sheet sourceSheet, Workbook targetWorkbook, String targetSheetName,
                                          int firstRow, int lastRow, int firstCol, int lastCol) {
        if (sourceSheet == null) {
            return null;
        }
        Sheet targetSheet = targetWorkbook.createSheet(targetSheetName);
        Map<Short, CellStyle> styleCache = new HashMap<>();
        for (int col = firstCol; col <= lastCol; col++) {
            int targetCol = col - firstCol;
            targetSheet.setColumnWidth(targetCol, sourceSheet.getColumnWidth(col));
            targetSheet.setColumnHidden(targetCol, sourceSheet.isColumnHidden(col));
        }
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row sourceRow = sourceSheet.getRow(rowIndex);
            Row targetRow = targetSheet.createRow(rowIndex - firstRow);
            if (sourceRow == null) {
                continue;
            }
            targetRow.setHeight(sourceRow.getHeight());
            targetRow.setZeroHeight(sourceRow.getZeroHeight());
            for (int col = firstCol; col <= lastCol; col++) {
                Cell sourceCell = sourceRow.getCell(col);
                if (sourceCell == null) {
                    continue;
                }
                Cell targetCell = targetRow.createCell(col - firstCol, sourceCell.getCellType());
                copyCellValue(sourceCell, targetCell);
                targetCell.setCellStyle(cloneCellStyle(targetWorkbook, sourceCell.getCellStyle(), styleCache));
            }
        }
        copyMergedRegions(sourceSheet, targetSheet, firstRow, lastRow, firstCol, lastCol);
        return targetSheet;
    }

    private static void setupMeasurementSheet(Sheet sheet) {
        if (sheet == null) {
            return;
        }
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
    }

    private static void setupMedColumnWidths(Sheet sheet) {
        if (sheet != null) {
            adjustColumnWidthPixels(sheet, 8, 4);
            adjustColumnWidthPixels(sheet, 9, -10);
            sheet.setColumnWidth(12, (int) Math.round((120d - 5d) / 7d * 256d));
            sheet.setMargin(Sheet.LeftMargin, 0.35d);
            sheet.setMargin(Sheet.RightMargin, 0.35d);
            sheet.setFitToPage(true);
            sheet.setAutobreaks(true);
            PrintSetup printSetup = sheet.getPrintSetup();
            printSetup.setFitWidth((short) 1);
            printSetup.setFitHeight((short) 0);
        }
    }

    private static void setupPprSheet(Sheet sheet) {
        if (sheet == null) {
            return;
        }
        sheet.setMargin(Sheet.LeftMargin, 0.35d);
        sheet.setMargin(Sheet.RightMargin, 0.35d);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
    }

    private static void adjustColumnWidthPixels(Sheet sheet, int column, int pixels) {
        int width = sheet.getColumnWidth(column);
        int delta = (int) Math.round(pixels / 7d * 256d);
        sheet.setColumnWidth(column, Math.max(256, width + delta));
    }

    private static int fillMedRows(Sheet sheet, Sheet sourceSheet, AreaProtocolPanel.AreaProtocolData data) {
        int profileCount = data.getGammaProfileMaxNumber();
        int controlPointCount = data.getGammaControlPointCount();
        if (sheet == null) {
            return 0;
        }
        if (profileCount + controlPointCount <= 0) {
            return 6;
        }
        Map<Short, CellStyle> styleCache = new HashMap<>();
        Row prototypeRow = sourceSheet == null ? null : sourceSheet.getRow(6);
        String medDate = findMedDate(data);
        for (int index = 0; index < profileCount; index++) {
            int rowIndex = 6 + index;
            Row row = createMedDataRow(sheet, prototypeRow, styleCache, rowIndex);
            row.getCell(0).setCellValue(index + 1);
            row.getCell(1).setCellValue("Профиль " + (index + 1));
            fillMedResultCells(sheet, row, prototypeRow, index + 1, data);
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 7));
        }
        int firstControlPointRow = 6 + profileCount;
        for (int index = 0; index < controlPointCount; index++) {
            int rowIndex = firstControlPointRow + index;
            Row row = createMedDataRow(sheet, prototypeRow, styleCache, rowIndex);
            row.getCell(0).setCellValue(profileCount + index + 1);
            row.getCell(1).setCellValue("Контрольная точка " + (index + 1));
            fillMedControlPointCells(sheet, row, prototypeRow, data);
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 7));
        }
        int totalRows = profileCount + controlPointCount;
        mergeMedColumnByPages(sheet, 6, totalRows, 8, medDate);
        mergeMedColumnByPages(sheet, 6, totalRows, 12, "0.30");
        applyFullBordersToColumnRows(sheet, 6, 5 + totalRows, 8);
        applyFullBordersToColumnRows(sheet, 6, 5 + totalRows, 12);
        return 6 + totalRows;
    }

    private static void mergeMedColumnByPages(Sheet sheet, int firstRow, int rowCount, int column, String value) {
        mergeColumnByPages(sheet, firstRow, rowCount, column, value, 12);
    }

    private static void mergePprColumnByPages(Sheet sheet, int firstRow, int rowCount, int column, String value) {
        mergeColumnByPages(sheet, firstRow, rowCount, column, value, 13);
    }

    private static void mergeColumnByPages(Sheet sheet, int firstRow, int rowCount, int column,
                                           String value, int lastColumn) {
        if (sheet == null || rowCount <= 0) {
            return;
        }
        double pageHeight = (595.28d - (sheet.getMargin(Sheet.TopMargin) + sheet.getMargin(Sheet.BottomMargin)) * 72d)
                / calculateFitWidthScale(sheet, lastColumn);
        double pageBreakTolerance = rowHeight(sheet, firstRow);
        double firstPageAvailable = pageHeight - rowsHeight(sheet, 0, firstRow - 1) + pageBreakTolerance;
        double nextPageAvailable = pageHeight - rowHeight(sheet, firstRow - 1) + pageBreakTolerance;
        int lastRow = firstRow + rowCount - 1;
        int blockStart = firstRow;
        double remaining = firstPageAvailable;
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            double height = rowHeight(sheet, rowIndex);
            if (rowIndex > blockStart && height > remaining) {
                mergeMedColumnBlock(sheet, blockStart, rowIndex - 1, column, value);
                blockStart = rowIndex;
                remaining = nextPageAvailable;
            }
            remaining -= height;
        }
        mergeMedColumnBlock(sheet, blockStart, lastRow, column, value);
    }

    private static double calculateFitWidthScale(Sheet sheet, int lastColumn) {
        double printableWidth = 841.89d - (sheet.getMargin(Sheet.LeftMargin) + sheet.getMargin(Sheet.RightMargin)) * 72d;
        double contentWidth = 0d;
        for (int column = 0; column <= lastColumn; column++) {
            contentWidth += columnWidthToPoints(sheet.getColumnWidth(column));
        }
        if (contentWidth <= 0d) {
            return 1d;
        }
        return Math.min(1d, printableWidth / contentWidth);
    }

    private static double columnWidthToPoints(int width) {
        return ((width / 256d) * 7d + 5d) * 0.75d;
    }

    private static void mergeMedColumnBlock(Sheet sheet, int firstRow, int lastRow, int column, String value) {
        Cell cell = getCell(sheet, new CellReference(firstRow, column).formatAsString());
        cell.setCellValue(value == null ? "" : value);
        if (lastRow > firstRow) {
            sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, column, column));
        }
    }

    private static double rowsHeight(Sheet sheet, int firstRow, int lastRow) {
        double height = 0d;
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            height += rowHeight(sheet, rowIndex);
        }
        return height;
    }

    private static double rowHeight(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row == null ? sheet.getDefaultRowHeightInPoints() : row.getHeightInPoints();
    }

    private static void applyFullBordersToColumnRows(Sheet sheet, int firstRow, int lastRow, int column) {
        Map<Short, CellStyle> styleCache = new HashMap<>();
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell cell = row.getCell(column);
            if (cell == null) {
                cell = row.createCell(column);
            }
            cell.setCellStyle(fullBorderStyle(sheet.getWorkbook(), cell.getCellStyle(), styleCache));
        }
    }

    private static CellStyle fullBorderStyle(Workbook workbook, CellStyle sourceStyle,
                                             Map<Short, CellStyle> styleCache) {
        short sourceIndex = sourceStyle.getIndex();
        CellStyle cachedStyle = styleCache.get(sourceIndex);
        if (cachedStyle != null) {
            return cachedStyle;
        }
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(sourceStyle);
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        styleCache.put(sourceIndex, style);
        return style;
    }

    private static int fillPprRows(Sheet sheet, Sheet sourceSheet, AreaProtocolPanel.AreaProtocolData data) {
        int pointCount = data.getPprPointCount();
        if (sheet == null) {
            return 0;
        }
        if (pointCount <= 0) {
            return 4;
        }
        Map<Short, CellStyle> styleCache = new HashMap<>();
        Row prototypeRow = sourceSheet == null ? null : sourceSheet.getRow(4);
        String pprDate = findPprDate(data);
        List<Double> pprValues = new ArrayList<>();
        for (int index = 0; index < pointCount; index++) {
            int rowIndex = 4 + index;
            Row row = createPprDataRow(sheet, prototypeRow, styleCache, rowIndex);
            row.getCell(0).setCellValue(index + 1);
            row.getCell(1).setCellValue("Точка " + (index + 1));
            Double value = fillPprResultCells(row, data);
            if (value != null) {
                pprValues.add(value);
            }
            sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 9));
        }
        mergePprColumnByPages(sheet, 4, pointCount, 10, pprDate);
        mergePprColumnByPages(sheet, 4, pointCount, 13, "80");
        int summaryRowIndex = 4 + pointCount;
        return addPprSummaryRow(sheet, prototypeRow, styleCache, summaryRowIndex, pprValues)
                ? summaryRowIndex + 1
                : summaryRowIndex;
    }

    private static Row createPprDataRow(Sheet sheet, Row prototypeRow, Map<Short, CellStyle> styleCache,
                                        int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        if (prototypeRow != null) {
            row.setHeight(prototypeRow.getHeight());
        }
        for (int col = 0; col <= 13; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellValue("");
            Cell prototypeCell = prototypeRow == null ? null : prototypeRow.getCell(col);
            if (prototypeCell != null) {
                cell.setCellStyle(cloneCellStyle(sheet.getWorkbook(), prototypeCell.getCellStyle(), styleCache));
            }
        }
        return row;
    }

    private static Double fillPprResultCells(Row row, AreaProtocolPanel.AreaProtocolData data) {
        double value = Math.round(randomRangeValue(
                parseDecimal(data.getPprMinValue()),
                parseDecimal(data.getPprMaxValue())
        ));
        if (value <= 0d) {
            return null;
        }
        row.getCell(11).setCellValue(value);
        row.getCell(12).setCellValue(Math.round(calculatePprUncertainty(value)));
        return value;
    }

    private static boolean addPprSummaryRow(Sheet sheet, Row prototypeRow, Map<Short, CellStyle> styleCache,
                                            int rowIndex, List<Double> values) {
        if (values.isEmpty()) {
            return false;
        }
        Row row = createPprDataRow(sheet, prototypeRow, styleCache, rowIndex);
        double average = average(values);
        double uncertainty = calculatePprSummaryUncertainty(values, average);
        Cell cell = row.getCell(0);
        cell.setCellValue(createPprSummaryText(sheet.getWorkbook(), Math.round(average), Math.round(uncertainty)));
        applyPprSummaryStyle(row);
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, 13));
        return true;
    }

    private static void appendAreaProtocolFooter(Sheet sheet, Sheet sourcePprSheet, int firstRow,
                                                 AreaProtocolPanel.AreaProtocolData data) {
        if (sheet == null) {
            return;
        }
        int footerLastColumn = data.isPprSelected() ? 13 : 12;
        if (sourcePprSheet != null && data.isPprSelected()) {
            sheet.setColumnWidth(13, sourcePprSheet.getColumnWidth(13));
            sheet.setColumnHidden(13, sourcePprSheet.isColumnHidden(13));
        }
        int customerRow = findRowContaining(sourcePprSheet, "Данные, предоставленные заказчиком");
        if (customerRow < 0) {
            customerRow = 90;
        }
        int[] sourceOffsets = {0, 1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13};
        double[] heights = {44.25d, 6d, 39.75d, 10.5d, 12d, 12.75d, 3.75d, 14.25d,
                12.75d, 4.5d, 12.75d, 12.75d, 12.75d, 26.25d, 19.5d, 12.75d};
        Map<Short, CellStyle> styleCache = new HashMap<>();
        for (int index = 0; index < sourceOffsets.length; index++) {
            Row row = createFooterRow(sheet, sourcePprSheet, customerRow + sourceOffsets[index],
                    firstRow + index, footerLastColumn, styleCache);
            row.setHeightInPoints((float) heights[index]);
        }
        getCell(sheet, new CellReference(firstRow, 0).formatAsString()).setCellValue(
                "U - значение расширенной неопределенности результата измерения, выраженное в абсолютных значениях. "
                        + "Указанная расширенная неопределенность измерений установлена как стандартная "
                        + "неопределенность измерений, умноженная на коэффициент охвата k (k=2), "
                        + "который соответствует вероятности охвата около 95 %");
        getCell(sheet, new CellReference(firstRow + 2, 0).formatAsString()).setCellValue(
                buildCustomerProvidedText(data));
        getCell(sheet, new CellReference(firstRow + 4, 0).formatAsString()).setCellValue(
                "Измерения проводил: заведующий лабораторией ____________ Тарновский М.О.");
        getCell(sheet, new CellReference(firstRow + 5, 0).formatAsString()).setCellValue(
                "(должность, подпись, Ф.И.О.)");
        getCell(sheet, new CellReference(firstRow + 7, 0).formatAsString()).setCellValue(
                "Протокол подготовил: заведующий лабораторией ____________ Тарновский М.О.");
        getCell(sheet, new CellReference(firstRow + 8, 0).formatAsString()).setCellValue(
                "(должность, подпись, Ф.И.О.)");
        getCell(sheet, new CellReference(firstRow + 10, 0).formatAsString()).setCellValue(
                "Испытательная лаборатория несет ответственность за всю информацию, представленную в протоколе испытаний, "
                        + "за исключением случаев, когда информация предоставляется заказчиком.\n"
                        + "Протокол не должен быть воспроизведен не в полном объеме без разрешения испытательной лаборатории ООО «ЦИТ».\n"
                        + "Результаты относятся только к объектам, прошедшим испытания и предоставленным заказчиком.\n"
                        + "Распределение экземпляров протокола испытаний: два протокола – Заказчику, один протокол – испытательной лаборатории ООО «ЦИТ».");
        getCell(sheet, new CellReference(firstRow + 15, 0).formatAsString()).setCellValue(
                "Конец протокола испытаний");
        sheet.addMergedRegion(new CellRangeAddress(firstRow, firstRow, 0, footerLastColumn));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 2, firstRow + 2, 0, footerLastColumn));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 4, firstRow + 4, 0, 12));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 5, firstRow + 5, 0, 12));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 7, firstRow + 7, 0, 12));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 8, firstRow + 8, 0, 12));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 10, firstRow + 14, 0, footerLastColumn));
        sheet.addMergedRegion(new CellRangeAddress(firstRow + 15, firstRow + 15, 0, footerLastColumn));
        applyLeftAlignment(sheet, firstRow + 4, 0, 12);
        applyLeftAlignment(sheet, firstRow + 7, 0, 12);
    }

    private static void applyPprSummaryStyle(Row row) {
        Map<Short, CellStyle> styleCache = new HashMap<>();
        for (int col = 0; col <= 13; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellStyle(pprSummaryStyle(row.getSheet().getWorkbook(), cell.getCellStyle(), styleCache));
        }
    }

    private static CellStyle pprSummaryStyle(Workbook workbook, CellStyle sourceStyle,
                                             Map<Short, CellStyle> styleCache) {
        short sourceIndex = sourceStyle.getIndex();
        CellStyle cachedStyle = styleCache.get(sourceIndex);
        if (cachedStyle != null) {
            return cachedStyle;
        }
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(sourceStyle);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        styleCache.put(sourceIndex, style);
        return style;
    }

    private static void applyLeftAlignment(Sheet sheet, int rowIndex, int firstCol, int lastCol) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return;
        }
        Map<Short, CellStyle> styleCache = new HashMap<>();
        for (int col = firstCol; col <= lastCol; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellStyle(leftAlignedStyle(sheet.getWorkbook(), cell.getCellStyle(), styleCache));
        }
    }

    private static CellStyle leftAlignedStyle(Workbook workbook, CellStyle sourceStyle,
                                              Map<Short, CellStyle> styleCache) {
        short sourceIndex = sourceStyle.getIndex();
        CellStyle cachedStyle = styleCache.get(sourceIndex);
        if (cachedStyle != null) {
            return cachedStyle;
        }
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(sourceStyle);
        style.setAlignment(HorizontalAlignment.LEFT);
        styleCache.put(sourceIndex, style);
        return style;
    }

    private static Row createFooterRow(Sheet targetSheet, Sheet sourceSheet, int sourceRowIndex, int targetRowIndex,
                                       int lastColumn, Map<Short, CellStyle> styleCache) {
        Row row = targetSheet.getRow(targetRowIndex);
        if (row == null) {
            row = targetSheet.createRow(targetRowIndex);
        }
        Row sourceRow = sourceSheet == null ? null : sourceSheet.getRow(sourceRowIndex);
        if (sourceRow != null) {
            row.setHeight(sourceRow.getHeight());
        }
        Cell sourceFirstCell = sourceRow == null ? null : sourceRow.getCell(0);
        for (int col = 0; col <= lastColumn; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellValue("");
            Cell sourceCell = sourceRow == null ? null : sourceRow.getCell(col);
            if (sourceCell == null) {
                sourceCell = sourceFirstCell;
            }
            if (sourceCell != null) {
                cell.setCellStyle(cloneCellStyle(targetSheet.getWorkbook(), sourceCell.getCellStyle(), styleCache));
            }
        }
        return row;
    }

    private static int findRowContaining(Sheet sheet, String text) {
        if (sheet == null || text == null || text.isBlank()) {
            return -1;
        }
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (formatter.formatCellValue(cell).contains(text)) {
                    return row.getRowNum();
                }
            }
        }
        return -1;
    }

    private static String buildCustomerProvidedText(AreaProtocolPanel.AreaProtocolData data) {
        String area = nullToEmpty(data.getAreaText()).isBlank() ? "____" : nullToEmpty(data.getAreaText());
        return "Данные, предоставленные заказчиком, за которые он несет ответственность:  Площадь общая "
                + area
                + " кв. м. Допустимый уровень в пункте 15. Испытательная лаборатория не несет ответственность "
                + "за данные, предоставленные заказчиком (заявление об ограничении ответственности испытательной лаборатории).";
    }

    private static double calculatePprSummaryUncertainty(List<Double> values, double average) {
        if (values.size() <= 1) {
            return 0d;
        }
        double sumSquares = 0d;
        for (double value : values) {
            double delta = value - average;
            sumSquares += delta * delta;
        }
        int pointCount = values.size();
        return Math.sqrt(sumSquares / pointCount * (pointCount - 1d));
    }

    private static RichTextString createPprSummaryText(Workbook workbook, long average, long uncertainty) {
        String text = "Среднее значение плотности потока радона с поверхности почвы с учетом неопределенности  Rср ± δ - ("
                + average + " ± " + uncertainty + ") мБк/(м2∙с) ";
        RichTextString richText = workbook.getCreationHelper().createRichTextString(text);
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        richText.applyFont(0, text.length(), baseFont);

        Font subFont = workbook.createFont();
        subFont.setFontName("Arial");
        subFont.setTypeOffset(Font.SS_SUB);
        int subIndex = text.indexOf("ср");
        if (subIndex >= 0) {
            richText.applyFont(subIndex, subIndex + 2, subFont);
        }

        Font superFont = workbook.createFont();
        superFont.setFontName("Arial");
        superFont.setTypeOffset(Font.SS_SUPER);
        int superIndex = text.indexOf("м2");
        if (superIndex >= 0) {
            richText.applyFont(superIndex + 1, superIndex + 2, superFont);
        }
        return richText;
    }

    private static double calculatePprUncertainty(double value) {
        if (value <= 0d) {
            return 0d;
        }
        return Math.sqrt(Math.pow(value * 0.3d / value, 2d) / 3d) * 2d * value;
    }

    private static Row createMedDataRow(Sheet sheet, Row prototypeRow, Map<Short, CellStyle> styleCache,
                                        int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        if (prototypeRow != null) {
            row.setHeight(prototypeRow.getHeight());
        }
        for (int col = 0; col <= 12; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }
            cell.setCellValue("");
            Cell prototypeCell = prototypeRow == null ? null : prototypeRow.getCell(col);
            if (prototypeCell != null) {
                cell.setCellStyle(cloneCellStyle(sheet.getWorkbook(), prototypeCell.getCellStyle(), styleCache));
            }
        }
        return row;
    }

    private static void fillMedResultCells(Sheet sheet, Row row, Row prototypeRow,
                                           int profileNumber, AreaProtocolPanel.AreaProtocolData data) {
        AreaProtocolPanel.MedValueData profile = findMedValue(data, profileNumber);
        int count = profile == null ? 0 : parsePositiveInt(profile.getCount());
        if (count <= 0 && profile != null) {
            count = calculateMedCount(profile.getDistance());
        }
        List<Double> values = createMedMeasurements(
                parseDecimal(data.getMedMinValue()),
                parseDecimal(data.getMedMaxValue()),
                parseDecimal(data.getMedAverageValue()),
                count
        );
        if (values.isEmpty()) {
            return;
        }
        double average = average(values);
        double sumSquares = sumSquares(values, average);
        double typeA = values.size() <= 1 || average == 0d
                ? 0d
                : (1d / average) * Math.sqrt(sumSquares / (values.size() * (values.size() - 1d))) * 100d;
        double typeB = average == 0d ? 0d : Math.sqrt(Math.pow(4d / average + 15d, 2d) / 3d);
        double combined = Math.sqrt(typeA * typeA + typeB * typeB);
        double expandedPercent = combined * 2d;
        double expandedValue = expandedPercent * average / 100d;
        String measurements = formatMeasurementList(values);
        row.getCell(9).setCellValue(measurements);
        row.getCell(10).setCellValue(round2(average));
        row.getCell(11).setCellValue(round2(expandedValue));
        adjustMedRowHeight(sheet, row, prototypeRow, measurements);
    }

    private static void fillMedControlPointCells(Sheet sheet, Row row, Row prototypeRow,
                                                 AreaProtocolPanel.AreaProtocolData data) {
        double value = round2(randomMedValue(
                parseDecimal(data.getMedMinValue()),
                parseDecimal(data.getMedMaxValue()),
                parseDecimal(data.getMedAverageValue())
        ));
        if (value <= 0d) {
            return;
        }
        row.getCell(9).setCellValue(value);
        row.getCell(10).setCellValue("-");
        row.getCell(11).setCellValue(round2(calculateMedControlPointUncertainty(value)));
        adjustMedRowHeight(sheet, row, prototypeRow, formatDecimal(value));
    }

    private static void adjustMedRowHeight(Sheet sheet, Row row, Row prototypeRow, String value) {
        if (sheet == null || row == null || prototypeRow == null) {
            return;
        }
        int lineCount = estimateWrappedLineCount(value, sheet.getColumnWidth(9));
        float lineHeight = prototypeRow.getHeightInPoints() / 2f;
        row.setHeightInPoints(Math.max(lineHeight, lineHeight * lineCount));
    }

    private static int estimateWrappedLineCount(String value, int columnWidth) {
        if (value == null || value.isBlank()) {
            return 1;
        }
        int charsPerLine = Math.max(1, (int) Math.floor(columnWidth / 256d));
        int lines = 0;
        for (String part : value.split("\\R", -1)) {
            lines += Math.max(1, (int) Math.ceil(part.length() / (double) charsPerLine));
        }
        return lines;
    }

    private static double calculateMedControlPointUncertainty(double value) {
        if (value <= 0d) {
            return 0d;
        }
        double ar = 15d + 4d / value;
        double as = value * ar / 100d;
        return Math.sqrt(Math.pow(as / value, 2d) / 3d) * 2d * value;
    }

    private static AreaProtocolPanel.MedValueData findMedValue(AreaProtocolPanel.AreaProtocolData data,
                                                               int profileNumber) {
        for (AreaProtocolPanel.MedValueData value : data.getMedValues()) {
            if (value.getProfileNumber() == profileNumber) {
                return value;
            }
        }
        return null;
    }

    private static List<Double> createMedMeasurements(double min, double max, double average, int count) {
        if (count <= 0 || min <= 0d || max <= 0d) {
            return List.of();
        }
        if (min > max) {
            double temp = min;
            min = max;
            max = temp;
        }
        if (average <= 0d) {
            average = (min + max) / 2d;
        }
        average = Math.max(min, Math.min(max, average));
        List<Double> values = new ArrayList<>();
        for (int i = 0; i < count; i++) {
            values.add(randomMedValue(min, max, average));
        }
        return values;
    }

    private static double randomMedValue(double min, double max, double targetAverage) {
        if (max <= min) {
            return min;
        }
        ThreadLocalRandom random = ThreadLocalRandom.current();
        double lowerMean = (min + targetAverage) / 2d;
        double upperMean = (targetAverage + max) / 2d;
        double lowerProbability = upperMean == lowerMean
                ? 1d
                : (upperMean - targetAverage) / (upperMean - lowerMean);
        lowerProbability = Math.max(0d, Math.min(1d, lowerProbability));
        if (random.nextDouble() < lowerProbability) {
            if (targetAverage <= min) {
                return min;
            }
            return random.nextDouble(min, targetAverage);
        }
        if (targetAverage >= max) {
            return max;
        }
        return random.nextDouble(targetAverage, max);
    }

    private static double randomRangeValue(double min, double max) {
        if (min <= 0d || max <= 0d) {
            return 0d;
        }
        if (min > max) {
            double temp = min;
            min = max;
            max = temp;
        }
        if (max <= min) {
            return min;
        }
        return ThreadLocalRandom.current().nextDouble(min, max);
    }

    private static double sum(List<Double> values) {
        double sum = 0d;
        for (double value : values) {
            sum += value;
        }
        return sum;
    }

    private static double average(List<Double> values) {
        return values.isEmpty() ? 0d : sum(values) / values.size();
    }

    private static double sumSquares(List<Double> values, double average) {
        double sum = 0d;
        for (double value : values) {
            double delta = value - average;
            sum += delta * delta;
        }
        return sum;
    }

    private static String formatMeasurementList(List<Double> values) {
        List<String> parts = new ArrayList<>();
        for (double value : values) {
            parts.add(formatDecimal(round2(value)));
        }
        return String.join("; ", parts);
    }

    private static int parsePositiveInt(String value) {
        if (value == null || value.isBlank()) {
            return 0;
        }
        try {
            return Math.max(0, Integer.parseInt(value.trim()));
        } catch (NumberFormatException ex) {
            return 0;
        }
    }

    private static int calculateMedCount(String distanceText) {
        double distance = parseDecimal(distanceText);
        if (distance <= 0d) {
            return 0;
        }
        double metersPerMeasurement = 5000d / 3600d * 8d;
        return Math.max(1, (int) Math.ceil(distance / metersPerMeasurement));
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

    private static double round2(double value) {
        return Math.round((value + 0.000000001d) * 100d) / 100d;
    }

    private static String formatDecimal(double value) {
        return String.format(Locale.ROOT, "%.2f", value).replace(".", ",");
    }

    private static String findMedDate(AreaProtocolPanel.AreaProtocolData data) {
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (row.isGammaSelected() && !nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (!nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        return nullToEmpty(data.getProtocolDate());
    }

    private static String findPprDate(AreaProtocolPanel.AreaProtocolData data) {
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (row.isPprSelected() && !nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        for (AreaProtocolPanel.MeasurementRowData row : data.getMeasurementRows()) {
            if (!nullToEmpty(row.getDate()).isBlank()) {
                return row.getDate();
            }
        }
        return nullToEmpty(data.getProtocolDate());
    }

    private static void superscriptSquareMeters(Sheet sheet) {
        if (sheet == null) {
            return;
        }
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() != CellType.STRING) {
                    continue;
                }
                String value = cell.getStringCellValue();
                if (value.contains("Rср")) {
                    continue;
                }
                if (value.contains("м2")) {
                    cell.setCellValue(value.replace("м2", "м²"));
                }
            }
        }
    }

    private static void copyCellValue(Cell sourceCell, Cell targetCell) {
        CellType type = sourceCell.getCellType();
        if (type == CellType.FORMULA) {
            targetCell.setCellFormula(sourceCell.getCellFormula());
            return;
        }
        if (type == CellType.NUMERIC) {
            targetCell.setCellValue(sourceCell.getNumericCellValue());
            return;
        }
        if (type == CellType.STRING) {
            targetCell.setCellValue(sourceCell.getStringCellValue());
            return;
        }
        if (type == CellType.BOOLEAN) {
            targetCell.setCellValue(sourceCell.getBooleanCellValue());
            return;
        }
        if (type == CellType.ERROR) {
            targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
        }
    }

    private static CellStyle cloneCellStyle(Workbook targetWorkbook, CellStyle sourceStyle,
                                            Map<Short, CellStyle> styleCache) {
        short sourceIndex = sourceStyle.getIndex();
        CellStyle cachedStyle = styleCache.get(sourceIndex);
        if (cachedStyle != null) {
            return cachedStyle;
        }
        CellStyle targetStyle = targetWorkbook.createCellStyle();
        targetStyle.cloneStyleFrom(sourceStyle);
        targetStyle.setDataFormat(targetWorkbook.createDataFormat().getFormat(sourceStyle.getDataFormatString()));
        styleCache.put(sourceIndex, targetStyle);
        return targetStyle;
    }

    private static void copyMergedRegions(Sheet sourceSheet, Sheet targetSheet,
                                          int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int index = 0; index < sourceSheet.getNumMergedRegions(); index++) {
            CellRangeAddress region = sourceSheet.getMergedRegion(index);
            if (region.getFirstRow() < firstRow || region.getLastRow() > lastRow
                    || region.getFirstColumn() < firstCol || region.getLastColumn() > lastCol) {
                continue;
            }
            targetSheet.addMergedRegion(new CellRangeAddress(
                    region.getFirstRow() - firstRow,
                    region.getLastRow() - firstRow,
                    region.getFirstColumn() - firstCol,
                    region.getLastColumn() - firstCol
            ));
        }
    }

    private static String buildProtocolNumber(AreaProtocolPanel.AreaProtocolData data) {
        String applicationNumber = nullToEmpty(data.getApplicationNumber());
        if (applicationNumber.isBlank()) {
            applicationNumber = "____";
        }
        return applicationNumber + "-1-Р/" + twoDigitYear(data);
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

    private static String buildReasonText(AreaProtocolPanel.AreaProtocolData data) {
        return "Основание для измерений: договор № " + nullToEmpty(data.getContractNumber())
                + " от " + nullToEmpty(data.getContractDate())
                + ", заявка № " + nullToEmpty(data.getApplicationNumber())
                + " от " + nullToEmpty(data.getApplicationDate());
    }

    private static String buildIndicatorsText(AreaProtocolPanel.AreaProtocolData data) {
        List<String> parts = new ArrayList<>();
        if (data.isPprSelected()) {
            parts.add(PPR_INDICATOR);
        }
        if (data.isMedSelected()) {
            parts.add(GAMMA_INDICATOR);
        }
        return parts.isEmpty() ? "-" : String.join(", ", parts);
    }

    private static String buildNdIndicatorText(AreaProtocolPanel.AreaProtocolData data) {
        if (data.isPprSelected() && data.isMedSelected()) {
            return PPR_ND_INDICATOR + "; " + GAMMA_INDICATOR;
        }
        if (data.isPprSelected()) {
            return PPR_ND_INDICATOR;
        }
        if (data.isMedSelected()) {
            return GAMMA_ND_INDICATOR;
        }
        return "-";
    }

    private static String buildAdditionalInfoText(AreaProtocolPanel.AreaProtocolData data) {
        String dates = buildDatesText(data.getMeasurementRows());
        StringBuilder text = new StringBuilder();
        text.append("Дополнительные сведения (характеристика объекта): Измерения были проведены ")
                .append(dates.isBlank() ? "____" : dates)
                .append(". Площадь общая ")
                .append(nullToEmpty(data.getAreaText()).isBlank() ? "____" : nullToEmpty(data.getAreaText()))
                .append(" кв. м. ");
        if (data.isPprSelected()) {
            text.append("Измерения ППР выполнены на участках открытого и сухого грунта. ");
        }
        text.append(buildWeatherText(data.getMeasurementRows()));
        text.append(" Эскиз предоставлен заказчиком. При измерениях на объекте заказчик самостоятельно определяет расположение точки в соответствии с эскизом.");
        return text.toString();
    }

    private static String buildDatesText(List<AreaProtocolPanel.MeasurementRowData> rows) {
        List<String> dates = new ArrayList<>();
        if (rows != null) {
            for (AreaProtocolPanel.MeasurementRowData row : rows) {
                if (row != null && row.getDate() != null && !row.getDate().isBlank() && !dates.contains(row.getDate())) {
                    dates.add(row.getDate());
                }
            }
        }
        return String.join(", ", dates);
    }

    private static String buildWeatherText(List<AreaProtocolPanel.MeasurementRowData> rows) {
        List<String> parts = new ArrayList<>();
        if (rows != null) {
            for (AreaProtocolPanel.MeasurementRowData row : rows) {
                if (row == null || row.getDate() == null || row.getDate().isBlank()) {
                    continue;
                }
                parts.add(row.getDate() + " составила " + buildTemperatureText(row)
                        + ", относительная влажность воздуха ____ %, атмосферное давление ____ мм рт. ст., без осадков");
            }
        }
        if (parts.isEmpty()) {
            return "Температура воздуха: ____ °С, относительная влажность воздуха ____ %, атмосферное давление ____ мм рт. ст., без осадков.";
        }
        return "Температура воздуха: " + String.join("; ", parts) + ".";
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
        return "____ °С";
    }

    private static String buildGammaSurveyText(AreaProtocolPanel.AreaProtocolData data) {
        return "Гамма-съемка территории проведена в режиме свободного поиска по маршрутным профилям, расстояние между которыми составляет не более "
                + gammaProfileDistance(data)
                + " на участке. Диапазон показаний дозиметра-радиометра: от " + formatMedRangeValue(data.getMedMinValue())
                + " мкЗв/ч до " + formatMedRangeValue(data.getMedMaxValue())
                + " мкЗв/ч. Радиационных аномалий на обследуемой территории  не обнаружено (в соответствии с МР 2.6.1.0361-24)";
    }

    private static String formatMedRangeValue(String value) {
        double parsed = parseDecimal(value);
        return parsed <= 0d ? "____" : formatDecimal(round2(parsed));
    }

    private static String gammaProfileDistance(AreaProtocolPanel.AreaProtocolData data) {
        double areaSquareMeters = parseAreaSquareMeters(data.getAreaText());
        if (areaSquareMeters <= 0) {
            return "2,5 м";
        }
        double hectares = areaSquareMeters / 10_000d;
        if (hectares <= 1d) {
            return "2,5 м";
        }
        if (hectares <= 5d) {
            return "5 м";
        }
        return "10 м";
    }

    private static double parseAreaSquareMeters(String value) {
        if (value == null) {
            return 0d;
        }
        String normalized = value.toLowerCase(Locale.ROOT)
                .replace('\u00A0', ' ')
                .replace("кв. м", "")
                .replace("м2", "")
                .replace("м²", "")
                .replaceAll("[^0-9,.-]", "")
                .replace(",", ".");
        if (normalized.isBlank()) {
            return 0d;
        }
        try {
            return Double.parseDouble(normalized);
        } catch (NumberFormatException ex) {
            return 0d;
        }
    }

    private static void setSuperscriptMinusSix(Sheet sheet) {
        Cell cell = getCell(sheet, "N40");
        String text = "Основная абсолютная погрешность, при температуре 25 ± 5 (˚С):\n"
                + "±(9,6·10-6 ·Тx+0,01) с\n"
                + "Дополнительная абсолютная погрешность при отклонении температуры от нормальных условий 25 ± 5 (˚С) на 1 ˚С изменения температуры:\n"
                + "-(2,2·10-6·Тx) с";
        RichTextString richText = sheet.getWorkbook().getCreationHelper().createRichTextString(text);
        Font baseFont = sheet.getWorkbook().createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 9);
        richText.applyFont(0, text.length(), baseFont);

        Font superFont = sheet.getWorkbook().createFont();
        superFont.setFontName("Arial");
        superFont.setFontHeightInPoints((short) 9);
        superFont.setTypeOffset(Font.SS_SUPER);
        int index = text.indexOf("-6");
        while (index >= 0) {
            richText.applyFont(index, index + 2, superFont);
            index = text.indexOf("-6", index + 2);
        }
        cell.setCellValue(richText);
    }

    private static void setCellText(Sheet sheet, String address, String value) {
        getCell(sheet, address).setCellValue(value == null ? "" : value);
    }

    private static Cell getCell(Sheet sheet, String address) {
        CellReference ref = new CellReference(address);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) {
            row = sheet.createRow(ref.getRow());
        }
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) {
            cell = row.createCell(ref.getCol());
        }
        return cell;
    }

    private static void highlightRange(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        for (int rowIndex = firstRow; rowIndex <= lastRow; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
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
        chooser.setDialogTitle("Сохранить протокол участка");
        chooser.setSelectedFile(new File("Протокол_участок.xlsx"));
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
