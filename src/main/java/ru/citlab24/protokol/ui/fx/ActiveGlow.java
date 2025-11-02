package ru.citlab24.protokol.ui.fx;

import javafx.animation.KeyFrame;
import javafx.animation.KeyValue;
import javafx.animation.Timeline;
import javafx.scene.control.Button;
import javafx.scene.effect.DropShadow;
import javafx.scene.paint.Color;
import javafx.util.Duration;

public final class ActiveGlow {
    private ActiveGlow() {}

    /** Неоновое свечение с пониженной яркостью (intensity ~ 0.5 = 50%). */
    public static void setActive(Button b, String hex, double intensity) {
        clear(b);

        Color c = Color.web(hex);

        DropShadow outer = new DropShadow();
        outer.setColor(c);
        outer.setRadius(0);
        outer.setSpread(0.35 * intensity);

        DropShadow outer2 = new DropShadow();
        outer2.setColor(c);
        outer2.setRadius(0);
        outer2.setSpread(0.15 * intensity);

        outer.setInput(outer2);
        b.setEffect(outer);

        Timeline t = new Timeline(
                new KeyFrame(Duration.ZERO,
                        new KeyValue(outer.radiusProperty(), 0),
                        new KeyValue(outer2.radiusProperty(), 0)),
                new KeyFrame(Duration.millis(160),
                        new KeyValue(outer.radiusProperty(), 25 * intensity),
                        new KeyValue(outer2.radiusProperty(), 55 * intensity))
        );
        t.play();
    }

    public static void clear(Button b) {
        if (b != null) b.setEffect(null);
    }
}
