package ru.citlab24.protokol.tabs.titleTab;

public record TitlePageImportData(
        String protocolDate,
        String customerNameAndContacts,
        String customerLegalAddress,
        String customerActualAddress,
        String objectName,
        String objectAddress,
        String contractNumber,
        String contractDate,
        String applicationNumber,
        String applicationDate
) {
}
