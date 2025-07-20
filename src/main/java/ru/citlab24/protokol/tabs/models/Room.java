package ru.citlab24.protokol.tabs.models;

public class Room {
    private String name = "";
    private Double volume = null;
    private int ventilationChannels = 1;
    private double ventilationSectionArea = 0.008;

    // Геттеры и сеттеры
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
        this.volume = volume;
    }
}