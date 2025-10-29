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
        // форма по умолчанию — квадрат; ширина — 0.100 м; сечение рассчитываем
        this(
                floor, space, room,
                Math.max(0, channels),
                area(DuctShape.SQUARE, 0.100),  // ВНИМАНИЕ: сечение ОДНОГО канала
                volume,
                roomRef, sectionIndex,
                DuctShape.SQUARE, 0.100
        );
    }

    public VentilationRecord withVolume(Double v)        { return new VentilationRecord(floor, space, room, channels, sectionArea, v, roomRef, sectionIndex, shape, width); }
    public VentilationRecord withShape(DuctShape s)      { return recalc(s, width); }
    public VentilationRecord withWidth(double w)         { return recalc(shape, w); }

    public VentilationRecord withChannels(int v) {
        // Кол-во каналов НЕ меняет «сечение» одного канала
        return new VentilationRecord(
                floor, space, room,
                Math.max(0, v),
                sectionArea,          // оставляем как было: площадь ОДНОГО канала
                volume, roomRef, sectionIndex,
                shape, width
        );
    }

    private VentilationRecord recalc(DuctShape s, double w) {
        // Сечение ВСЕГДА на один канал
        double aOne = area(s, w);
        return new VentilationRecord(
                floor, space, room,
                channels,
                aOne,                // площадь одного канала
                volume, roomRef, sectionIndex,
                s, w
        );
    }

    // Площадь одного канала по форме
    public static double area(DuctShape s, double w) {
        if (w <= 0) return 0.0;
        return switch (s) {
            case SQUARE -> w * w;
            case CIRCLE -> Math.PI * Math.pow(w / 2.0, 2.0);
        };
    }

}
