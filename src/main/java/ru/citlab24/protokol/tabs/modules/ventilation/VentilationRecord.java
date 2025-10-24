package ru.citlab24.protokol.tabs.modules.ventilation;

import ru.citlab24.protokol.tabs.models.Room;

public record VentilationRecord(
        String floor,
        String space,
        String room,
        int channels,
        double sectionArea,   // вычисляется
        Double volume,
        Room roomRef,
        int sectionIndex,
        DuctShape shape,      // новая колонка
        double width          // новая колонка, м; для круга — диаметр
) {
    public enum DuctShape { SQUARE { public String toString(){return "квадрат";} }, CIRCLE { public String toString(){return "круг";} } }

    public VentilationRecord(String floor, String space, String room,
                             int channels, double sectionArea, Double volume,
                             Room roomRef, int sectionIndex) {
        this(floor, space, room, channels, sectionArea, volume, roomRef, sectionIndex,
                DuctShape.SQUARE,
                sectionArea > 0 ? Math.sqrt(sectionArea) : 0.1 // дефолт 0.1 → 0.1×0.1≈0.01
        );
    }

    public VentilationRecord withVolume(Double v)        { return new VentilationRecord(floor, space, room, channels, sectionArea, v, roomRef, sectionIndex, shape, width); }
    public VentilationRecord withChannels(int v)         { return new VentilationRecord(floor, space, room, v, sectionArea, volume, roomRef, sectionIndex, shape, width); }
    public VentilationRecord withShape(DuctShape s)      { return recalc(s, width); }
    public VentilationRecord withWidth(double w)         { return recalc(shape, w); }

    public static double area(DuctShape s, double w) {
        if (w <= 0) return 0.0;
        return switch (s) {
            case SQUARE -> w * w;
            case CIRCLE -> Math.PI * Math.pow(w / 2.0, 2.0);
        };
    }
    private VentilationRecord recalc(DuctShape s, double w) {
        double a = area(s, w);
        return new VentilationRecord(floor, space, room, channels, a, volume, roomRef, sectionIndex, s, w);
    }
}
