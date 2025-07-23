package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Floor {
    private int id;
    private String number;
    private String name;
    private FloorType type;
    private List<Space> spaces = new ArrayList<>();

    public enum FloorType {
        RESIDENTIAL("жилой"),
        OFFICE("офисный"),
        PUBLIC("общественный"),
        MIXED("смешанный");

        private final String title;

        FloorType(String title) {
            this.title = title;
        }

        @Override
        public String toString() {
            return title;
        }
    }

    // Геттеры и сеттеры
    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public void addSpace(Space space) {
        spaces.add(space);
    }

    public String getNumber() {
        return number;
    }

    public FloorType getType() {
        return type;
    }

    public void setType(FloorType type) {
        this.type = type;
    }

    public void setNumber(String number) {
        this.number = number;
    }

    public void setSpaces(List<Space> spaces) {
        this.spaces = spaces;
    }

    public List<Space> getSpaces() {
        return spaces;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }
}
