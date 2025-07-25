package ru.citlab24.protokol.tabs.models;

import java.util.ArrayList;
import java.util.List;

public class Space {
    private String identifier;
    private SpaceType type;
    private List<Room> rooms = new ArrayList<>();

    public enum SpaceType {
        APARTMENT("Квартира"),
        OFFICE("Офис"),
        PUBLIC_SPACE("Общественный");    // Общественное пространство

        // Другое

        private final String title;

        SpaceType(String title) {
            this.title = title;
        }

        @Override
        public String toString() {
            return title;
        }
    }


    // Геттеры и сеттеры
    public void addRoom(Room room) {
        rooms.add(room);
    }
    public String getIdentifier() {
        return identifier;
    }
    public void setIdentifier(String identifier) {
        this.identifier = identifier;
    }
    public SpaceType getType() {
        return type;
    }
    public void setType(SpaceType type) {
        this.type = type;
    }
    public List<Room> getRooms() {
        return rooms;
    }
}
