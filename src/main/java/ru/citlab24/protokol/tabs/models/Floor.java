package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Floor {
    private int id;
    private String number;
    private String name;
    private FloorType type;
    private int sectionIndex = 0;
    private int position = 0;

    public int getPosition() { return position; }
    public void setPosition(int position) { this.position = Math.max(0, position); }

    private List<Space> spaces = new ArrayList<>();

    public enum FloorType {
        RESIDENTIAL("жилой"),
        OFFICE("офисный"),
        PUBLIC("общественный"),
        MIXED("смешанный"),
        STREET("улица"); // новый тип этажа «УЛИЦА»

        public final String title;
        FloorType(String title) { this.title = title; }
        @Override public String toString() { return title; }
    }

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getNumber() { return number; }
    public void setNumber(String number) { this.number = number; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public FloorType getType() { return type; }
    public void setType(FloorType type) { this.type = type; }

    public int getSectionIndex() { return sectionIndex; }
    public void setSectionIndex(int sectionIndex) { this.sectionIndex = Math.max(0, sectionIndex); }

    public List<Space> getSpaces() { return spaces; }

    public void addSpace(Space space) { spaces.add(space); }
}
