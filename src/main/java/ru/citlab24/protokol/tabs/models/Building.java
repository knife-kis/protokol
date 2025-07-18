package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Building {
    private int id;
    private String name;
    private List<Floor> floors = new ArrayList<>();

    // Геттеры и сеттеры
    public void addFloor(Floor floor) {
        floors.add(floor);
    }
    public List<Floor> getFloors() {
        return floors;
    }
    public String getName() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }

    public void setFloors(List<Floor> floors) {
        this.floors = floors;
    }

    public int getId() {
        return id;
    }

    public void setId(int id) {
        this.id = id;
    }
}
