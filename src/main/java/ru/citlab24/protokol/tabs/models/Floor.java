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

        public final String title;

        FloorType(String title) {
            this.title = title;
        }

        @Override
        public String toString() {
            return title;
        }
    }

    @Override
    public String toString() {
        // Если есть имя этажа - используем его
        if (name != null && !name.trim().isEmpty()) {
            return name;
        }

        // Если есть номер этажа - используем его
        if (number != null && !number.trim().isEmpty()) {
            return number;
        }

        // Если тип этажа задан - используем его
        if (type != null) {
            return type.title;
        }

        // Все остальные случаи
        return "Этаж";
    }

    // Геттеры и сеттеры
    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getNumber() { return number; }
    public void setNumber(String number) { this.number = number; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public FloorType getType() { return type; }
    public void setType(FloorType type) { this.type = type; }

    public List<Space> getSpaces() { return spaces; }
    public void setSpaces(List<Space> spaces) { this.spaces = spaces; }

    public void addSpace(Space space) {
        spaces.add(space);
    }
}
