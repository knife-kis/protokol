package ru.citlab24.protokol.export;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageSetup;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDuplex;

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
        pageSetup.setDuplex(STDuplex.SHORT);
    }
}
