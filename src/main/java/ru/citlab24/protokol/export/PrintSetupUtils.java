package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;

import javax.xml.namespace.QName;

public final class PrintSetupUtils {
    private PrintSetupUtils() {}

    public static final double DEFAULT_MARGIN_LEFT_CM = 0.8;
    public static final double DEFAULT_MARGIN_RIGHT_CM = 0.5;
    public static final double DEFAULT_MARGIN_TOP_CM = 3.3;
    public static final double DEFAULT_MARGIN_BOTTOM_CM = 1.9;

    public static void applyDuplexShortEdge(Sheet sheet) {
        if (!(sheet instanceof XSSFSheet)) {
            return;
        }
        XSSFSheet xssfSheet = (XSSFSheet) sheet;

        CTPageSetup pageSetup = xssfSheet.getCTWorksheet().isSetPageSetup()
                ? xssfSheet.getCTWorksheet().getPageSetup()
                : xssfSheet.getCTWorksheet().addNewPageSetup();

        // Ставим атрибут duplex напрямую, без STDuplex (которого может не быть в схемах)
        XmlCursor cursor = null;
        try {
            cursor = pageSetup.newCursor();
            cursor.toStartDoc();
            cursor.toFirstChild(); // <pageSetup .../>
            cursor.setAttributeText(new QName("", "duplex"), "duplexShortEdge");
        } catch (Throwable ignore) {
            // не валим экспорт
        } finally {
            if (cursor != null) {
                cursor.dispose();
            }
        }
    }

    public static void applyDuplexShortEdge(Workbook workbook) {
        if (workbook == null) {
            return;
        }
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            if (sheet != null) {
                applyDuplexShortEdge(sheet);
            }
        }
    }

    public static void applyStandardMargins(Sheet sheet) {
        if (sheet == null) {
            return;
        }
        sheet.setMargin(Sheet.LeftMargin, cmToInches(DEFAULT_MARGIN_LEFT_CM));
        sheet.setMargin(Sheet.RightMargin, cmToInches(DEFAULT_MARGIN_RIGHT_CM));
        sheet.setMargin(Sheet.TopMargin, cmToInches(DEFAULT_MARGIN_TOP_CM));
        sheet.setMargin(Sheet.BottomMargin, cmToInches(DEFAULT_MARGIN_BOTTOM_CM));
    }

    private static double cmToInches(double cm) {
        return cm / 2.54;
    }
}
