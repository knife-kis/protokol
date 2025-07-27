package ru.citlab24.protokol.tabs.utils;

import java.util.List;
import java.util.Locale;

public class RoomUtils {
    public static final List<String> RESIDENTIAL_ROOM_KEYWORDS = List.of(
            "кухня", "кухня-ниша", "кухня-гостиная", "кухня гостиная", "кухня ниша",
            "санузел", "сан узел", "сан. узел",
            "ванная", "ванная комната",
            "совмещенный", "совмещенный санузел", "туалет",
            "мусорокамера", "су", "с/у"
    );

    public static boolean isResidentialRoom(String roomName) {
        if (roomName == null) return false;
        String normalized = normalizeRoomName(roomName);
        return RESIDENTIAL_ROOM_KEYWORDS.stream().anyMatch(normalized::contains);
    }

    public static Double getAirExchangeRate(String roomName) {
        if (roomName == null) return null;

        String normalized = normalizeRoomName(roomName);

        if (normalized.contains("кухня") || normalized.contains("кухня-ниша") || normalized.contains("кухня ниша") || normalized.contains("кухня-гостиная") || normalized.contains("кухня гостиная")) {
            return 60.0;
        } else if (normalized.contains("ванная комната") ||
                normalized.contains("совмещенный санузел") ||
                normalized.contains("ванная")) {
            return 50.0;
        } else if (normalized.contains("санузел") ||
                normalized.contains("сан узел") ||
                normalized.contains("сан. узел") ||
                normalized.contains("туалет")) {
            return 25.0;
        }
        return null;
    }

    public static String normalizeRoomName(String roomName) {
        return roomName
                .replaceAll("[\\s\\.-]+", " ")
                .trim()
                .toLowerCase(Locale.ROOT);
    }
}
