package ru.citlab24.protokol.db;

import java.util.ArrayList;
import java.util.List;

public class PersonnelRecord {
    private int id;
    private String firstName;
    private String lastName;
    private String middleName;
    private final List<UnavailabilityRecord> unavailabilityDates = new ArrayList<>();

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getFirstName() {
        return firstName;
    }

    public void setFirstName(String firstName) {
        this.firstName = firstName;
    }

    public String getLastName() {
        return lastName;
    }

    public void setLastName(String lastName) {
        this.lastName = lastName;
    }

    public String getMiddleName() {
        return middleName;
    }

    public void setMiddleName(String middleName) {
        this.middleName = middleName;
    }

    public List<UnavailabilityRecord> getUnavailabilityDates() {
        return unavailabilityDates;
    }

    public String getFullName() {
        return (safe(lastName) + " " + safe(firstName) + " " + safe(middleName)).trim();
    }

    private String safe(String value) {
        return value == null ? "" : value.trim();
    }

    public static class UnavailabilityRecord {
        private int id;
        private String unavailableDate;
        private String reason;

        public int getId() {
            return id;
        }

        public void setId(int id) {
            this.id = id;
        }

        public String getUnavailableDate() {
            return unavailableDate;
        }

        public void setUnavailableDate(String unavailableDate) {
            this.unavailableDate = unavailableDate;
        }

        public String getReason() {
            return reason;
        }

        public void setReason(String reason) {
            this.reason = reason;
        }
    }
}
