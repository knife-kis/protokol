package ru.citlab24.protokol.db;

public class VlkDateRecord {
    private int id;
    private String vlkDate;
    private String responsible;
    private String eventName;

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getVlkDate() {
        return vlkDate;
    }

    public void setVlkDate(String vlkDate) {
        this.vlkDate = vlkDate;
    }

    public String getResponsible() {
        return responsible;
    }

    public void setResponsible(String responsible) {
        this.responsible = responsible;
    }

    public String getEventName() {
        return eventName;
    }

    public void setEventName(String eventName) {
        this.eventName = eventName;
    }
}
