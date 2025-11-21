package ru.citlab24.protokol.tabs.models;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Building implements Serializable {
    private static final long serialVersionUID = 1L;
    private int plannedFloorsCount;
    private int id;
    private String name;
    private final List<Section> sections = new ArrayList<>(); // <-- секции
    private final List<Floor> floors = new ArrayList<>();

    public int getId() { return id; }
    public void setId(int id) { this.id = id; }

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public List<Floor> getFloors() { return floors; }
    public void addFloor(Floor floor) { floors.add(floor); }


    // секции
    public List<Section> getSections() { return sections; }
    public void setSections(List<Section> newSections) {
        sections.clear();
        if (newSections != null) sections.addAll(newSections);
    }
    public void addSection(Section s) { sections.add(s); }
}
