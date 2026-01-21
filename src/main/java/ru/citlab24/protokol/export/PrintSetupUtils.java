package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;

import javax.xml.namespace.QName;

public final class PrintSetupUtils {
    private PrintSetupUtils() {}

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
}
