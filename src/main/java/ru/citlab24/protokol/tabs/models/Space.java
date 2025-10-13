package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Space {
    private static int nextId = 1;

    private int id;
    private String identifier;
    private SpaceType type;
    private List<Room> rooms = new ArrayList<>();
    private int position = 0;
    public int getPosition() { return position; }
    public void setPosition(int position) { this.position = Math.max(0, position); }

    public Space() {
        this.id = nextId++; // ← уникальный ID для каждого помещения
    }

    public enum SpaceType {
        APARTMENT("Квартира"),
        OFFICE("Офис"),
        PUBLIC_SPACE("Общественный");

        private final String title;

        SpaceType(String title) {
            this.title = title;
        }

        @Override
        public String toString() {
            return title;
        }
    }

    public int getId() { return id; }
    public void setId(int id) { this.id = id; } // оставим, если где-то явно задаёшь
    @Override
    public String toString() { return identifier + " (" + type + ")"; }

    public void addRoom(Room room) { rooms.add(room); }
    public String getIdentifier() { return identifier; }
    public void setIdentifier(String identifier) { this.identifier = identifier; }
    public SpaceType getType() { return type; }
    public void setType(SpaceType type) { this.type = type; }
    public List<Room> getRooms() { return rooms; }
}
