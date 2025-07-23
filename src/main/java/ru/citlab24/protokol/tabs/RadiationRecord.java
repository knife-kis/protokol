package ru.citlab24.protokol.tabs;

public class RadiationRecord {
    private String building;
    private String floor;
    private String room;
    private Double measurementResult; // Результат измерения
    private Double permissibleLevel;   // Допустимый уровень

    public RadiationRecord() {}

    // Геттеры и сеттеры
    public String getFloor() { return floor; }
    public void setFloor(String floor) { this.floor = floor; }

    public String getRoom() { return room; }
    public void setRoom(String room) { this.room = room; }

    public Double getMeasurementResult() { return measurementResult; }
    public void setMeasurementResult(Double measurementResult) { this.measurementResult = measurementResult; }

    public Double getPermissibleLevel() { return permissibleLevel; }
    public void setPermissibleLevel(Double permissibleLevel) { this.permissibleLevel = permissibleLevel; }
    public String getBuilding() { return building; }
    public void setBuilding(String building) { this.building = building; }
}