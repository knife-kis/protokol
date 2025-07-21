package ru.citlab24.protokol.tabs;

import ru.citlab24.protokol.tabs.models.Room;

public record VentilationRecord(
        String floor,
        String space,
        String room,
        int channels,
        double sectionArea,
        Double volume,
        Room roomRef
) {
    public VentilationRecord withVolume(Double newVolume) {
        return new VentilationRecord(floor, space, room, channels, sectionArea, newVolume, roomRef);
    }
    public VentilationRecord withChannels(int newChannels) {
        return new VentilationRecord(floor, space, room, newChannels, sectionArea, volume, roomRef);
    }
    public VentilationRecord withSectionArea(double newArea) {
        return new VentilationRecord(floor, space, room, channels, newArea, volume, roomRef);
    }
}