package ru.citlab24.protokol.tabs.modules.med;

public class RadiationRecord {
    private String building;
    private String floor;
    private String space; // Добавлено поле для помещения
    private String room;
    private Double measurementResult;
    private Double permissibleLevel;

    public RadiationRecord() {}

    // Геттеры и сеттеры
    public String getBuilding() { return building; }
    public void setBuilding(String building) { this.building = building; }

    public String getFloor() { return floor; }
    public void setFloor(String floor) { this.floor = floor; }

    public String getSpace() { return space; }
    public void setSpace(String space) { this.space = space; }

    public String getRoom() { return room; }
    public void setRoom(String room) { this.room = room; }

    public Double getMeasurementResult() { return measurementResult; }
    public void setMeasurementResult(Double measurementResult) { this.measurementResult = measurementResult; }

    public Double getPermissibleLevel() { return permissibleLevel; }
    public void setPermissibleLevel(Double permissibleLevel) { this.permissibleLevel = permissibleLevel; }
}