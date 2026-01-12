package ru.citlab24.protokol.tabs.titleTab;

import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.StringJoiner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public final class TechnicalAssignmentImporter {
    private static final Pattern DATE_PATTERN = Pattern.compile("(\\d{2}\\.\\d{2}\\.\\d{4})");
    private static final Pattern CUSTOMER_LABEL = Pattern.compile("(?i)ЗАКАЗЧИК\\s*:?\\s*(.*)");
    private static final Pattern EMAIL_LABEL = Pattern.compile(
            "(?i)Адрес\\s+электронной\\s+почты\\s+Заказчика\\s*:?\\s*(.*)");
    private static final Pattern ADDRESS_LABEL = Pattern.compile("(?i)Адрес\\s*:?\\s*(.*)");
    private static final String ANCHOR_SECTION = "АДРЕСА, РЕКВИЗИТЫ И ПОДПИСИ СТОРОН";

    private TechnicalAssignmentImporter() {
    }

    public static TitlePageImportData importFromFile(File file) throws IOException {
        if (file == null) {
            throw new IOException("Файл не выбран.");
        }
        if (!file.getName().toLowerCase(Locale.ROOT).endsWith(".docx")) {
            throw new IOException("Поддерживается только формат .docx.");
        }

        try (FileInputStream inputStream = new FileInputStream(file);
             XWPFDocument document = new XWPFDocument(inputStream)) {
            List<String> lines = extractLines(document);
            String protocolDate = extractProtocolDate(lines);
            int anchorIndex = findAnchorIndex(lines);
            int searchStart = (anchorIndex >= 0) ? anchorIndex + 1 : 0;

            CustomerExtract customerExtract = extractCustomerName(lines, searchStart);
            String email = extractValueAfterLabel(lines, searchStart, EMAIL_LABEL);
            String customerNameAndContacts = buildNameAndContacts(customerExtract.name(), email);
            String address = extractAddress(lines, customerExtract.lineIndex(), searchStart);

            return new TitlePageImportData(
                    protocolDate,
                    customerNameAndContacts,
                    address,
                    address
            );
        }
    }

    private static List<String> extractLines(XWPFDocument document) {
        List<String> lines = new ArrayList<>();
        for (IBodyElement element : document.getBodyElements()) {
            if (element instanceof XWPFParagraph paragraph) {
                addParagraphLines(lines, paragraph.getText());
            } else if (element instanceof XWPFTable table) {
                addTableLines(lines, table);
            }
        }
        return lines;
    }

    private static void addParagraphLines(List<String> lines, String text) {
        if (text == null) {
            return;
        }
        String[] split = text.split("\\R");
        for (String line : split) {
            lines.add(line);
        }
    }

    private static void addTableLines(List<String> lines, XWPFTable table) {
        for (XWPFTableRow row : table.getRows()) {
            StringJoiner joiner = new StringJoiner(" ");
            for (XWPFTableCell cell : row.getTableCells()) {
                String cellText = cell.getText();
                if (cellText != null && !cellText.isBlank()) {
                    joiner.add(cellText.trim());
                }
            }
            String rowText = joiner.toString().trim();
            if (!rowText.isBlank()) {
                addParagraphLines(lines, rowText);
            }
        }
    }

    private static String extractProtocolDate(List<String> lines) {
        for (String line : lines) {
            if (line == null || line.isBlank()) {
                continue;
            }
            Matcher matcher = DATE_PATTERN.matcher(line);
            if (matcher.find()) {
                return matcher.group(1);
            }
            break;
        }
        return "";
    }

    private static int findAnchorIndex(List<String> lines) {
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            if (line.toUpperCase(Locale.ROOT).contains(ANCHOR_SECTION)) {
                return i;
            }
        }
        return -1;
    }

    private static CustomerExtract extractCustomerName(List<String> lines, int startIndex) {
        for (int i = startIndex; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            Matcher matcher = CUSTOMER_LABEL.matcher(line);
            if (matcher.find()) {
                List<String> parts = new ArrayList<>();
                String initial = matcher.group(1);
                if (initial != null && !initial.isBlank()) {
                    parts.add(initial);
                }
                for (int j = i + 1; j < lines.size(); j++) {
                    String next = lines.get(j);
                    if (next == null || next.isBlank()) {
                        break;
                    }
                    if (isCustomerNameStopLine(next)) {
                        break;
                    }
                    parts.add(next);
                }
                String name = normalizeSpace(String.join(" ", parts));
                return new CustomerExtract(name, i);
            }
        }
        return new CustomerExtract("", -1);
    }

    private static boolean isCustomerNameStopLine(String line) {
        String upper = line.toUpperCase(Locale.ROOT);
        return upper.contains("АДРЕС")
                || upper.contains("ЭЛЕКТРОННОЙ ПОЧТЫ")
                || upper.contains("ИСПОЛНИТЕЛЬ")
                || upper.contains("ПОДПИСИ")
                || upper.contains("РЕКВИЗИТ");
    }

    private static String extractValueAfterLabel(List<String> lines, int startIndex, Pattern label) {
        for (int i = startIndex; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            Matcher matcher = label.matcher(line);
            if (matcher.find()) {
                String value = matcher.group(1);
                if (value == null || value.isBlank()) {
                    if (i + 1 < lines.size()) {
                        value = lines.get(i + 1);
                    }
                }
                return normalizeSpace(value);
            }
        }
        return "";
    }

    private static String extractAddress(List<String> lines, int customerIndex, int startIndex) {
        int from = (customerIndex >= 0) ? customerIndex : startIndex;
        for (int i = from; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            if (line.toUpperCase(Locale.ROOT).contains("ЭЛЕКТРОННОЙ ПОЧТЫ")) {
                continue;
            }
            Matcher matcher = ADDRESS_LABEL.matcher(line);
            if (matcher.find()) {
                String value = matcher.group(1);
                if (value == null || value.isBlank()) {
                    if (i + 1 < lines.size()) {
                        value = lines.get(i + 1);
                    }
                }
                return normalizeSpace(value);
            }
        }
        return "";
    }

    private static String buildNameAndContacts(String name, String email) {
        String normalizedName = normalizeSpace(name);
        String normalizedEmail = normalizeSpace(email);
        if (!normalizedName.isBlank() && !normalizedEmail.isBlank()) {
            return normalizedName + ". " + normalizedEmail;
        }
        if (!normalizedName.isBlank()) {
            return normalizedName;
        }
        return normalizedEmail;
    }

    private static String normalizeSpace(String value) {
        if (value == null) {
            return "";
        }
        return value.replaceAll("\\s+", " ").trim();
    }

    private record CustomerExtract(String name, int lineIndex) {
    }
}
