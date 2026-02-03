package ru.citlab24.protokol.protocolmap.area;

import ru.citlab24.protokol.protocolmap.RequestFormExporter;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileOutputStream;

final class SamplingPlanExporter {
    private static final String PLAN_NAME = "план отбора.docx";
    private static final String FONT_NAME = "Arial";
    private static final int FONT_SIZE = 12;

    private SamplingPlanExporter() {
    }

    static void generate(File mapFile) {
        if (mapFile == null || !mapFile.exists()) {
            return;
        }
        File targetFile = resolveSamplingPlanFile(mapFile);
        String applicationNumber = RequestFormExporter.resolveApplicationNumberFromMap(mapFile);

        try (XWPFDocument document = new XWPFDocument()) {
            RequestFormExporter.applyStandardHeader(document);

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("План отбора");
            titleRun.setFontFamily(FONT_NAME);
            titleRun.setFontSize(FONT_SIZE);
            titleRun.setBold(true);

            XWPFParagraph number = document.createParagraph();
            number.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun numberRun = number.createRun();
            numberRun.setText("Заявка № " + applicationNumber);
            numberRun.setFontFamily(FONT_NAME);
            numberRun.setFontSize(FONT_SIZE);

            XWPFParagraph spacer = document.createParagraph();
            spacer.createRun().addBreak();

            XWPFParagraph note = document.createParagraph();
            note.setAlignment(ParagraphAlignment.LEFT);
            XWPFRun noteRun = note.createRun();
            noteRun.setText("План отбора сформирован автоматически для первички по участкам.");
            noteRun.setFontFamily(FONT_NAME);
            noteRun.setFontSize(FONT_SIZE);

            try (FileOutputStream out = new FileOutputStream(targetFile)) {
                document.write(out);
            }
        } catch (Exception ignored) {
            // пропускаем создание листа, если не удалось сформировать документ
        }
    }

    static File resolveSamplingPlanFile(File mapFile) {
        if (mapFile == null) {
            return null;
        }
        return new File(mapFile.getParentFile(), PLAN_NAME);
    }
}
