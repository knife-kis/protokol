package ru.citlab24.protokol.tabs.models;

import java.io.Serializable;

public class Room implements Serializable {
    private static final long serialVersionUID = 1L;
    private Integer originalRoomId;
    private boolean selected;
    private static int nextId = 1;
    private int id;
    private String name = "";
    private Double volume = null;
    private int ventilationChannels = 1;
    private double ventilationSectionArea = 0.008;
    private Integer externalWallsCount = null; // Микроклимат: 0..4, null = не задано (проставим по умолчанию)

    private int position = 0;
    public int getPosition() { return position; }
    public void setPosition(int position) { this.position = Math.max(0, position); }
    public Room() {
        this.id = nextId++;
        this.originalRoomId = this.id;
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
    // --- Микроклимат: независимая галочка выбора (НЕ связана с освещением) ---
    private boolean microclimateSelected;
    public boolean isMicroclimateSelected() { return microclimateSelected; }
    public void setMicroclimateSelected(boolean microclimateSelected) { this.microclimateSelected = microclimateSelected; }
    // --- Радиация (ионизирующие излучения): независимая галочка выбора ---
    private boolean radiationSelected;
    public boolean isRadiationSelected() { return radiationSelected; }
    public void setRadiationSelected(boolean radiationSelected) { this.radiationSelected = radiationSelected; }



    public void setSelected(boolean selected) {
        this.selected = selected;
    }
    public Integer getExternalWallsCount() {
        return externalWallsCount;
    }
    public void setExternalWallsCount(Integer count) {
        if (count == null) {
            this.externalWallsCount = null;
            return;
        }
        int v = Math.max(0, Math.min(4, count));
        this.externalWallsCount = v;
    }

}