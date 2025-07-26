package ru.citlab24.protokol.tabs.models;

public class Room {
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

    public void setName(String name) {
        this.name = name;
    }

    public Double getVolume() {
        return volume;
    }

    public void setVolume(Double volume) {
        this.volume = (volume != null && volume == 0.0) ? null : volume;
    }

}