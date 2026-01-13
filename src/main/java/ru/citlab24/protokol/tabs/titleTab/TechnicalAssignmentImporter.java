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
    private static final Pattern TEXT_DATE_PATTERN = Pattern.compile(
            "«?\\s*(\\d{1,2})\\s*»?\\s*([А-Яа-я]+)\\s*(\\d{4})\\s*г?\\.?");
    private static final Pattern CUSTOMER_LABEL = Pattern.compile("(?i)ЗАКАЗЧИК\\s*:?\\s*(.*)");
    private static final Pattern EMAIL_LABEL = Pattern.compile(
            "(?i)Адрес\\s+электронной\\s+почты\\s+Заказчика\\s*:?\\s*(.*)");
    private static final Pattern CUSTOMER_ADDRESS_LABEL = Pattern.compile(
            "(?i)(?:юр\\s*/\\s*почтовый\\s+адрес|адрес)\\s*:?\\s*(.*)");
    private static final Pattern CUSTOMER_PHONE_LABEL = Pattern.compile("(?i)Тел\\.?\\s*:?\\s*(.*)");
    private static final Pattern OBJECT_NAME_LABEL = Pattern.compile(
            "(?i)Наименование\\s+объекта\\s+испытаний\\s*\\(исследований\\)");
    private static final Pattern OBJECT_ADDRESS_LABEL = Pattern.compile("(?i)Адрес\\s+объекта");
    private static final Pattern APPLICATION_LABEL = Pattern.compile(
            "(?i)Основание\\s+для\\s+проведения\\s+работ\\s*:?\\s*(.*)");
    private static final Pattern CONTRACT_NUMBER_PATTERN = Pattern.compile("(?i)ДОГОВОР\\s*№\\s*([^\\s]+)");
    private static final Pattern CONTRACT_DATE_HINT = Pattern.compile("(?i)ОТ\\s+");
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
            String email = extractValueAfterLabel(lines, 0, EMAIL_LABEL);
            String phone = extractCustomerSectionValue(lines, anchorIndex, CUSTOMER_PHONE_LABEL, 10);
            String customerNameAndContacts = buildNameAndContacts(customerExtract.name(), email, phone);
            String address = extractCustomerSectionValue(lines, anchorIndex, CUSTOMER_ADDRESS_LABEL, 10);
            String objectName = extractTableValue(document, OBJECT_NAME_LABEL);
            String objectAddress = extractTableValue(document, OBJECT_ADDRESS_LABEL);
            String contractNumber = extractContractNumber(lines);
            String contractDate = extractContractDate(lines);
            String applicationText = extractTableValue(document, APPLICATION_LABEL);
            if (applicationText.isBlank()) {
                applicationText = extractValueAfterLabel(lines, 0, APPLICATION_LABEL);
            }
            ApplicationExtract applicationExtract = extractApplicationInfo(applicationText);

            return new TitlePageImportData(
                    protocolDate,
                    customerNameAndContacts,
                    address,
                    address,
                    objectName,
                    objectAddress,
                    contractNumber,
                    contractDate,
                    applicationExtract.number(),
                    applicationExtract.date()
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

    private static String extractTableValue(XWPFDocument document, Pattern labelPattern) {
        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                List<XWPFTableCell> cells = row.getTableCells();
                for (int i = 0; i < cells.size(); i++) {
                    String cellText = normalizeSpace(cells.get(i).getText());
                    if (cellText.isBlank()) {
                        continue;
                    }
                    Matcher matcher = labelPattern.matcher(cellText);
                    if (!matcher.find()) {
                        continue;
                    }
                    String remainder = normalizeSpace(cellText.substring(matcher.end()));
                    if (!remainder.isBlank()) {
                        return remainder;
                    }
                    for (int j = i + 1; j < cells.size(); j++) {
                        String value = normalizeSpace(cells.get(j).getText());
                        if (!value.isBlank()) {
                            return value;
                        }
                    }
                }
            }
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
                String initial = matcher.group(1);
                return new CustomerExtract(normalizeSpace(initial), i);
            }
        }
        return new CustomerExtract("", -1);
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

    private static String extractCustomerSectionValue(List<String> lines, int anchorIndex, Pattern labelPattern,
                                                      int lookAhead) {
        int customerIndex = findCustomerSectionIndex(lines, anchorIndex);
        if (customerIndex < 0) {
            return "";
        }
        int start = Math.min(lines.size(), customerIndex + 1);
        int end = Math.min(lines.size(), start + lookAhead);
        for (int i = start; i < end; i++) {
            String line = lines.get(i);
            if (line == null || line.isBlank()) {
                continue;
            }
            if (line.toUpperCase(Locale.ROOT).contains("ЭЛЕКТРОННОЙ ПОЧТЫ")) {
                continue;
            }
            Matcher matcher = labelPattern.matcher(line);
            if (matcher.find()) {
                String value = matcher.group(1);
                if ((value == null || value.isBlank()) && i + 1 < end) {
                    value = lines.get(i + 1);
                }
                return normalizeSpace(value);
            }
        }
        return "";
    }

    private static int findCustomerSectionIndex(List<String> lines, int anchorIndex) {
        int start = anchorIndex >= 0 ? anchorIndex : 0;
        for (int i = start; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null) {
                continue;
            }
            if (line.toUpperCase(Locale.ROOT).contains("ЗАКАЗЧИК")) {
                return i;
            }
        }
        return -1;
    }

    private static String extractContractNumber(List<String> lines) {
        for (String line : lines) {
            if (line == null || line.isBlank()) {
                continue;
            }
            Matcher matcher = CONTRACT_NUMBER_PATTERN.matcher(line);
            if (matcher.find()) {
                return normalizeSpace(matcher.group(1));
            }
            if (line.toUpperCase(Locale.ROOT).contains("ДОГОВОР")) {
                return "";
            }
        }
        return "";
    }

    private static String extractContractDate(List<String> lines) {
        int contractIndex = -1;
        for (int i = 0; i < lines.size(); i++) {
            String line = lines.get(i);
            if (line == null || line.isBlank()) {
                continue;
            }
            if (line.toUpperCase(Locale.ROOT).contains("ДОГОВОР")) {
                contractIndex = i;
                String date = extractDateFromLine(line);
                if (!date.isBlank()) {
                    return date;
                }
                break;
            }
        }
        if (contractIndex >= 0) {
            int start = Math.min(lines.size(), contractIndex + 1);
            int end = Math.min(lines.size(), start + 8);
            for (int i = start; i < end; i++) {
                String line = lines.get(i);
                if (line == null || line.isBlank()) {
                    continue;
                }
                String date = extractDateFromLine(line);
                if (!date.isBlank()) {
                    return date;
                }
            }
        }
        int start = (contractIndex >= 0) ? contractIndex : 0;
        int end = Math.min(lines.size(), start + 6);
        for (int i = start; i < end; i++) {
            String line = lines.get(i);
            if (line == null || line.isBlank()) {
                continue;
            }
            if (CONTRACT_DATE_HINT.matcher(line).find()) {
                String date = extractDateFromLine(line);
                if (!date.isBlank()) {
                    return date;
                }
            }
        }
        return "";
    }

    private static ApplicationExtract extractApplicationInfo(String value) {
        if (value == null || value.isBlank()) {
            return new ApplicationExtract("", "");
        }
        String normalized = normalizeSpace(value);
        String number = "";
        String date = "";
        Matcher matcher = Pattern.compile("(?i)Заявка\\s*№\\s*([^\\s]+)").matcher(normalized);
        if (matcher.find()) {
            number = matcher.group(1);
            int stopIndex = number.toLowerCase(Locale.ROOT).indexOf("от");
            if (stopIndex >= 0) {
                number = number.substring(0, stopIndex).trim();
            }
        }
        int fromIndex = normalized.toLowerCase(Locale.ROOT).indexOf("от");
        if (fromIndex >= 0) {
            String after = normalized.substring(fromIndex + 2).trim();
            date = extractDateFromLine(after);
        } else {
            date = extractDateFromLine(normalized);
        }
        return new ApplicationExtract(normalizeSpace(number), date);
    }

    private static String extractDateFromLine(String line) {
        if (line == null || line.isBlank()) {
            return "";
        }
        Matcher numeric = DATE_PATTERN.matcher(line);
        if (numeric.find()) {
            return numeric.group(1);
        }
        Matcher text = TEXT_DATE_PATTERN.matcher(line);
        if (text.find()) {
            int day = Integer.parseInt(text.group(1));
            int month = resolveMonth(text.group(2));
            int year = Integer.parseInt(text.group(3));
            if (month > 0) {
                return String.format("%02d.%02d.%04d", day, month, year);
            }
        }
        return "";
    }

    private static int resolveMonth(String value) {
        if (value == null) {
            return 0;
        }
        String normalized = value.toLowerCase(Locale.ROOT);
        return switch (normalized) {
            case "января" -> 1;
            case "февраля" -> 2;
            case "марта" -> 3;
            case "апреля" -> 4;
            case "мая" -> 5;
            case "июня" -> 6;
            case "июля" -> 7;
            case "августа" -> 8;
            case "сентября" -> 9;
            case "октября" -> 10;
            case "ноября" -> 11;
            case "декабря" -> 12;
            default -> 0;
        };
    }

    private static String buildNameAndContacts(String name, String email, String phone) {
        String normalizedName = normalizeSpace(name);
        String normalizedEmail = normalizeSpace(email);
        String normalizedPhone = normalizeSpace(phone);
        List<String> contacts = new ArrayList<>();
        if (!normalizedEmail.isBlank()) {
            contacts.add(normalizedEmail);
        }
        if (!normalizedPhone.isBlank()) {
            contacts.add(normalizedPhone);
        }
        String contactsValue = String.join(", ", contacts);
        if (!normalizedName.isBlank() && !contactsValue.isBlank()) {
            return normalizedName + ". " + contactsValue;
        }
        if (!normalizedName.isBlank()) {
            return normalizedName;
        }
        return contactsValue;
    }

    private static String normalizeSpace(String value) {
        if (value == null) {
            return "";
        }
        return value.replaceAll("\\s+", " ").trim();
    }

    private record CustomerExtract(String name, int lineIndex) {
    }

    private record ApplicationExtract(String number, String date) {
    }
}
