package ru.citlab24.protokol.tabs.modules.ventilation;

import ru.citlab24.protokol.tabs.models.Room;

public record VentilationRecord(
        String floor,
        String space,
        String room,
        int channels,
        double sectionArea,
        Double volume,
        Room roomRef,
        int sectionIndex // 0-based индекс блок-секции (Floor.getSectionIndex())
) {
    public VentilationRecord withVolume(Double newVolume) {
        return new VentilationRecord(floor, space, room, channels, sectionArea, newVolume, roomRef, sectionIndex);
    }
    public VentilationRecord withChannels(int newChannels) {
        return new VentilationRecord(floor, space, room, newChannels, sectionArea, volume, roomRef, sectionIndex);
    }
    public VentilationRecord withSectionArea(double newArea) {
        return new VentilationRecord(floor, space, room, channels, newArea, volume, roomRef, sectionIndex);
    }
}


