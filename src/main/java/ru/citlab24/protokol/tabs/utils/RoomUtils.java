package ru.citlab24.protokol.tabs.utils;

import ru.citlab24.protokol.tabs.models.Floor;
import ru.citlab24.protokol.tabs.models.Room;

import java.util.List;
import java.util.Locale;

public class RoomUtils {
    public static final List<String> RESIDENTIAL_ROOM_KEYWORDS = List.of(
            "кухн", "кухня", "кухня-ниша", "кухня-гостиная", "кухня гостиная", "кухня ниша",
            "санузел", "сан узел", "сан. узел",
            "ванная", "ванная комната",
            "совмещенный", "совмещенный санузел", "туалет", "уборная",
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

        if (normalized.contains("кухн") || normalized.contains("кухня-ниша") ||
                normalized.contains("кухня ниша") || normalized.contains("кухня-гостиная") ||
                normalized.contains("кухня гостиная")) {
            return 60.0;
        } else if (normalized.contains("ванная комната") ||
                normalized.contains("совмещенный санузел") ||
                normalized.contains("ванная")) {
            return 50.0;
        } else if (normalized.contains("санузел") ||
                normalized.contains("сан узел") ||
                normalized.contains("сан. узел") ||
                normalized.contains("туалет") ||
                normalized.contains("уборная")) {   // ← добавили «уборная»
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
    /** Комнаты «мусор…»: мусорокамера, мусорная камера, мусоросборная и т. п. */
    public static boolean isPublicTrashRoom(String roomName) {
        if (roomName == null) return false;
        String normalized = normalizeRoomName(roomName);
        return normalized.contains("мусор"); // покрывает мусорокамеру/мусорные/мусоросборную
    }

    /**
     * Единый предикат для вкладки «Вентиляция»:
     *  - Жилые/смешанные: берём стандартные жилые санпомещения (кухни, с/у, ванны, туалет/уборная и т. п.)
     *  - Общественные: дополнительно берём все «мусор…» комнаты
     *  - Офисные: поведение оставляем как для жилых санпомещений (если нужно — можно расширить позже)
     */
    public static boolean isVentilationRelevant(Floor.FloorType floorType, String roomName) {
        if (roomName == null) return false;
        if (floorType == null) return isResidentialRoom(roomName);

        switch (floorType) {
            case PUBLIC:
                return isResidentialRoom(roomName) || isPublicTrashRoom(roomName);
            case MIXED:
                return isResidentialRoom(roomName) || isPublicTrashRoom(roomName);
            case RESIDENTIAL:
            case OFFICE:
            default:
                return isResidentialRoom(roomName);
        }
    }

}
