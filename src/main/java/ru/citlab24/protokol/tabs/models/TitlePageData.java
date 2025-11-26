package ru.citlab24.protokol.tabs.models;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

public class TitlePageData implements Serializable {
    private static final long serialVersionUID = 1L;

    private String protocolDate;
    private String customerNameAndContacts;
    private String customerLegalAddress;
    private String customerActualAddress;
    private String objectName;
    private String objectAddress;
    private String contractNumber;
    private String contractDate;
    private String applicationNumber;
    private String applicationDate;
    private String representative;
    private List<Measurement> measurements = new ArrayList<>();

    public String getProtocolDate() { return protocolDate; }
    public void setProtocolDate(String protocolDate) { this.protocolDate = protocolDate; }

    public String getCustomerNameAndContacts() { return customerNameAndContacts; }
    public void setCustomerNameAndContacts(String customerNameAndContacts) { this.customerNameAndContacts = customerNameAndContacts; }

    public String getCustomerLegalAddress() { return customerLegalAddress; }
    public void setCustomerLegalAddress(String customerLegalAddress) { this.customerLegalAddress = customerLegalAddress; }

    public String getCustomerActualAddress() { return customerActualAddress; }
    public void setCustomerActualAddress(String customerActualAddress) { this.customerActualAddress = customerActualAddress; }

    public String getObjectName() { return objectName; }
    public void setObjectName(String objectName) { this.objectName = objectName; }

    public String getObjectAddress() { return objectAddress; }
    public void setObjectAddress(String objectAddress) { this.objectAddress = objectAddress; }

    public String getContractNumber() { return contractNumber; }
    public void setContractNumber(String contractNumber) { this.contractNumber = contractNumber; }

    public String getContractDate() { return contractDate; }
    public void setContractDate(String contractDate) { this.contractDate = contractDate; }

    public String getApplicationNumber() { return applicationNumber; }
    public void setApplicationNumber(String applicationNumber) { this.applicationNumber = applicationNumber; }

    public String getApplicationDate() { return applicationDate; }
    public void setApplicationDate(String applicationDate) { this.applicationDate = applicationDate; }

    public String getRepresentative() { return representative; }
    public void setRepresentative(String representative) { this.representative = representative; }

    public List<Measurement> getMeasurements() { return measurements; }
    public void setMeasurements(List<Measurement> measurements) {
        this.measurements = (measurements == null)
                ? new ArrayList<>()
                : new ArrayList<>(measurements);
    }

    public TitlePageData copy() {
        TitlePageData copy = new TitlePageData();
        copy.setProtocolDate(getProtocolDate());
        copy.setCustomerNameAndContacts(getCustomerNameAndContacts());
        copy.setCustomerLegalAddress(getCustomerLegalAddress());
        copy.setCustomerActualAddress(getCustomerActualAddress());
        copy.setObjectName(getObjectName());
        copy.setObjectAddress(getObjectAddress());
        copy.setContractNumber(getContractNumber());
        copy.setContractDate(getContractDate());
        copy.setApplicationNumber(getApplicationNumber());
        copy.setApplicationDate(getApplicationDate());
        copy.setRepresentative(getRepresentative());

        List<Measurement> rowsCopy = new ArrayList<>();
        for (Measurement m : getMeasurements()) {
            if (m != null) {
                rowsCopy.add(m.copy());
            }
        }
        copy.setMeasurements(rowsCopy);
        return copy;
    }

    public static class Measurement implements Serializable {
        private static final long serialVersionUID = 1L;

        private String date;
        private String tempInsideStart;
        private String tempInsideEnd;
        private String tempOutsideStart;
        private String tempOutsideEnd;

        public String getDate() { return date; }
        public void setDate(String date) { this.date = date; }

        public String getTempInsideStart() { return tempInsideStart; }
        public void setTempInsideStart(String tempInsideStart) { this.tempInsideStart = tempInsideStart; }

        public String getTempInsideEnd() { return tempInsideEnd; }
        public void setTempInsideEnd(String tempInsideEnd) { this.tempInsideEnd = tempInsideEnd; }

        public String getTempOutsideStart() { return tempOutsideStart; }
        public void setTempOutsideStart(String tempOutsideStart) { this.tempOutsideStart = tempOutsideStart; }

        public String getTempOutsideEnd() { return tempOutsideEnd; }
        public void setTempOutsideEnd(String tempOutsideEnd) { this.tempOutsideEnd = tempOutsideEnd; }

        public Measurement copy() {
            Measurement copy = new Measurement();
            copy.setDate(getDate());
            copy.setTempInsideStart(getTempInsideStart());
            copy.setTempInsideEnd(getTempInsideEnd());
            copy.setTempOutsideStart(getTempOutsideStart());
            copy.setTempOutsideEnd(getTempOutsideEnd());
            return copy;
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (o == null || getClass() != o.getClass()) return false;
            Measurement that = (Measurement) o;
            return Objects.equals(date, that.date) &&
                    Objects.equals(tempInsideStart, that.tempInsideStart) &&
                    Objects.equals(tempInsideEnd, that.tempInsideEnd) &&
                    Objects.equals(tempOutsideStart, that.tempOutsideStart) &&
                    Objects.equals(tempOutsideEnd, that.tempOutsideEnd);
        }

        @Override
        public int hashCode() {
            return Objects.hash(date, tempInsideStart, tempInsideEnd, tempOutsideStart, tempOutsideEnd);
        }
    }
}
