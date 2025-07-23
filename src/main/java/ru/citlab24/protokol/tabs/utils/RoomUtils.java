package ru.citlab24.protokol.tabs.utils;

import java.util.List;
import java.util.Locale;

public class RoomUtils {
    public static final List<String> RESIDENTIAL_ROOM_KEYWORDS = List.of(
            "кухня", "кухня-ниша",
            "санузел", "сан узел", "сан. узел",
            "ванная", "ванная комната",
            "совмещенный", "совмещенный санузел", "туалет",
            "мусорокамера"
    );

    public static boolean isResidentialRoom(String roomName) {
        if (roomName == null) return false;

        String normalized = roomName
                .replaceAll("[\\s.-]+", " ")
                .trim()
                .toLowerCase(Locale.ROOT);

        return RESIDENTIAL_ROOM_KEYWORDS.stream()
                .anyMatch(normalized::contains);
    }
}
