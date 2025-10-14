package ru.citlab24.protokol.tabs.modules.lighting;

import ru.citlab24.protokol.tabs.models.Room;

public record LightingRecord(
        boolean selected,
        String floor,
        String space,
        String room,
        Room roomRef,
        int sectionIndex // чтобы знать, к какой секции относится
) {
    public LightingRecord withSelected(boolean newSelected) {
        return new LightingRecord(newSelected, floor, space, room, roomRef, sectionIndex);
    }
}
