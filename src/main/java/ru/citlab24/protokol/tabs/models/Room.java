package ru.citlab24.protokol.tabs.models;

public class Room {
    private Integer originalRoomId;
    private boolean selected;
    private static int nextId = 1;
    private int id;
    private String name = "";
    private Double volume = null;
    private int ventilationChannels = 1;
    private double ventilationSectionArea = 0.008;

    public Room() {
        this.id = nextId++;
    }

    // Геттеры и сеттеры
    public Integer getOriginalRoomId() {
        return originalRoomId;
    }

    public void setOriginalRoomId(Integer originalRoomId) {
        this.originalRoomId = originalRoomId;
    }
    public int getId() {
        return id;
    }
    public void setId(int id) {
        this.id = id;
    }
    public int getVentilationChannels() {
        return ventilationChannels;
    }
    public void setVentilationChannels(int ventilationChannels) {
        this.ventilationChannels = ventilationChannels;
    }
    public double getVentilationSectionArea() {
        return ventilationSectionArea;
    }
    public void setVentilationSectionArea(double ventilationSectionArea) {
        this.ventilationSectionArea = ventilationSectionArea;
    }
    public String getName() {
        return name;
    }

    @Override
    public String toString() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }
    public Double getVolume() {
        return volume;
    }
    public void setVolume(Double volume) {
        this.volume = (volume != null && volume == 0.0) ? null : volume;
    }
    public boolean isSelected() {
        return selected;
    }

    public void setSelected(boolean selected) {
        this.selected = selected;
    }
}