package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Building {
    private int plannedFloorsCount;
    private int id;
    private String name;
    private final List<Floor> floors = new ArrayList<>();

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<Floor> getFloors() {
        return floors;
    }

    public void addFloor(Floor floor) {
        floors.add(floor);
    }
    public int getPlannedFloorsCount() {
        return plannedFloorsCount;
    }

    public void setPlannedFloorsCount(int plannedFloorsCount) {
        this.plannedFloorsCount = plannedFloorsCount;
    }
}