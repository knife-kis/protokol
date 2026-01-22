package ru.citlab24.protokol.protocolmap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public final class PhysicalFactorsMapExporter {
    private static final String REGISTRATION_PREFIX = "Регистрационный номер карты замеров:";
    private static final double COLUMN_WIDTH_SCALE = 0.9;
    private static final double LEFT_MARGIN_CM = 0.8;
    private static final double RIGHT_MARGIN_CM = 0.5;
    private static final double TOP_MARGIN_CM = 3.3;
    private static final double BOTTOM_MARGIN_CM = 1.9;

    private PhysicalFactorsMapExporter() {
    }

    public static File generateMap(File sourceFile) throws IOException {
        String registrationNumber = resolveRegistrationNumber(sourceFile);
        File targetFile = buildTargetFile(sourceFile);

        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("карта замеров");
            applySheetDefaults(workbook, sheet);
            applyHeaders(sheet, registrationNumber);
            createTitleRow(workbook, sheet);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                workbook.write(out);
            }
        }

        return targetFile;
    }

    private static String resolveRegistrationNumber(File sourceFile) {
        if (sourceFile == null || !sourceFile.exists()) {
            return "";
        }
        try (InputStream in = new FileInputStream(sourceFile);
             Workbook workbook = WorkbookFactory.create(in)) {
            if (workbook.getNumberOfSheets() == 0) {
                return "";
            }
            Sheet sheet = workbook.getSheetAt(0);
            return findRegistrationNumber(sheet);
        } catch (Exception ex) {
            return "";
        }
    }

    private static String findRegistrationNumber(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            for (Cell cell : row) {
                String text = formatter.formatCellValue(cell).trim();
                if (text.startsWith(REGISTRATION_PREFIX)) {
                    String tail = text.substring(REGISTRATION_PREFIX.length()).trim();
                    if (!tail.isEmpty()) {
                        return tail;
                    }
                    Cell next = row.getCell(cell.getColumnIndex() + 1);
                    if (next != null) {
                        String nextText = formatter.formatCellValue(next).trim();
                        if (!nextText.isEmpty()) {
                            return nextText;
                        }
                    }
                }
            }
        }
        return "";
    }

    private static File buildTargetFile(File sourceFile) {
        String name = sourceFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = dotIndex > 0 ? name.substring(0, dotIndex) : name;
        return new File(sourceFile.getParentFile(), baseName + "_карта.xlsx");
    }

    private static void applySheetDefaults(Workbook workbook, Sheet sheet) {
        Font baseFont = workbook.createFont();
        baseFont.setFontName("Arial");
        baseFont.setFontHeightInPoints((short) 12);
        CellStyle baseStyle = workbook.createCellStyle();
        baseStyle.setFont(baseFont);

        int[] widthsPx = buildColumnWidthsPx();
        for (int col = 0; col < widthsPx.length; col++) {
            sheet.setColumnWidth(col, pixel2WidthUnits(widthsPx[col]));
            sheet.setDefaultColumnStyle(col, baseStyle);
        }

        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);
        printSetup.setFitWidth((short) 1);
        printSetup.setFitHeight((short) 0);
        sheet.setFitToPage(true);
        sheet.setAutobreaks(true);

        sheet.setMargin(Sheet.LeftMargin, cmToInches(LEFT_MARGIN_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(RIGHT_MARGIN_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(TOP_MARGIN_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(BOTTOM_MARGIN_CM));
    }

    private static void applyHeaders(Sheet sheet, String registrationNumber) {
        String font = "&\"Arial\"&12";
        Header header = sheet.getHeader();
        header.setLeft(font + "Испытательная лаборатория\nООО «ЦИТ»");
        header.setCenter(font + "Карта замеров № " + registrationNumber + "\nФ8 РИ ИЛ 2-2023");
        header.setRight(font + "\nКоличество страниц: &[Страница] / &[Страниц] \n ");
    }

    private static void createTitleRow(Workbook workbook, Sheet sheet) {
        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setFontHeightInPoints((short) 16);
        titleFont.setBold(true);

        CellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(org.apache.poi.ss.usermodel.HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(org.apache.poi.ss.usermodel.VerticalAlignment.CENTER);

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("Испытательная лаборатория Общества с ограниченной ответственностью");
        cell.setCellStyle(titleStyle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 31));
    }

    private static int[] buildColumnWidthsPx() {
        int[] baseWidths = new int[32];
        baseWidths[0] = 36;
        baseWidths[1] = 33;
        baseWidths[2] = 32;
        baseWidths[3] = 32;
        baseWidths[4] = 32;
        baseWidths[5] = 32;
        baseWidths[6] = 32;
        baseWidths[7] = 32;
        baseWidths[8] = 32;
        for (int col = 9; col <= 17; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[18] = 53;
        baseWidths[19] = 41;
        baseWidths[20] = 58;
        baseWidths[21] = 32;
        for (int col = 22; col <= 29; col++) {
            baseWidths[col] = 32;
        }
        baseWidths[30] = 36;
        baseWidths[31] = 11;

        int[] widths = new int[baseWidths.length];
        for (int i = 0; i < baseWidths.length; i++) {
            widths[i] = (int) Math.round(baseWidths[i] * COLUMN_WIDTH_SCALE);
        }
        return widths;
    }

    private static int pixel2WidthUnits(int px) {
        int units = (px / 7) * 256;
        int rem = px % 7;
        final int[] offset = {0, 36, 73, 109, 146, 182, 219};
        units += offset[rem];
        return units;
    }

    private static double cmToInches(double centimeters) {
        return centimeters / 2.54d;
    }
}
