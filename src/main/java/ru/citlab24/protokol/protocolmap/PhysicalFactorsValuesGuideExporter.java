package ru.citlab24.protokol.protocolmap;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public final class PhysicalFactorsValuesGuideExporter {
    private static final String GUIDE_NAME = "Справка по значениям.docx";

    private PhysicalFactorsValuesGuideExporter() {
    }

    public static File generate(File mapFile) throws IOException {
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
            addParagraph(document,
                    "В графе «средняя горизонтальная освещенность» заполняем два столбца:");
            addBullet(document,
                    "Первый столбец берём из файла, который подгружали для формирования первички: "
                            + "вкладка «Иск освещение (2)», столбец C, начиная с 8 строки.");
            addBullet(document,
                    "Второй столбец берём из той же вкладки «Иск освещение (2)»: "
                            + "ячейки K, L, M, N, …, AD (всего 20 значений, указываем через точку с запятой).");

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
}
