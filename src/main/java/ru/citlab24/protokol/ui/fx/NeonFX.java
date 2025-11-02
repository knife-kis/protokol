package ru.citlab24.protokol.ui.fx;

import javafx.animation.KeyFrame;
import javafx.animation.KeyValue;
import javafx.animation.Timeline;
import javafx.scene.control.Button;
import javafx.scene.effect.DropShadow;
import javafx.util.Duration;

/** Неоновая подсветка для JavaFX Button (hover: glow + лёгкий zoom). */
public final class NeonFX {
    private NeonFX() {}

    public static void apply(Button b, String hexColor) {
        final javafx.scene.paint.Color c = javafx.scene.paint.Color.web(hexColor);

        final DropShadow outer = new DropShadow(0, c);
        outer.setSpread(0.35);

        final DropShadow outer2 = new DropShadow(0, c);
        outer2.setSpread(0.15);
        outer.setInput(outer2);

        b.setOnMouseEntered(ev -> {
            b.setEffect(outer);
            Timeline t = new Timeline(
                    new KeyFrame(Duration.ZERO,
                            new KeyValue(outer.radiusProperty(), 0),
                            new KeyValue(outer2.radiusProperty(), 0),
                            new KeyValue(b.scaleXProperty(), 1.0),
                            new KeyValue(b.scaleYProperty(), 1.0)
                    ),
                    new KeyFrame(Duration.millis(180),
                            new KeyValue(outer.radiusProperty(), 35),
                            new KeyValue(outer2.radiusProperty(), 80),
                            new KeyValue(b.scaleXProperty(), 1.02),
                            new KeyValue(b.scaleYProperty(), 1.02)
                    )
            );
            t.play();
        });

        b.setOnMouseExited(ev -> {
            Timeline t = new Timeline(
                    new KeyFrame(Duration.millis(160),
                            new KeyValue(outer.radiusProperty(), 0),
                            new KeyValue(outer2.radiusProperty(), 0),
                            new KeyValue(b.scaleXProperty(), 1.0),
                            new KeyValue(b.scaleYProperty(), 1.0)
                    )
            );
            t.setOnFinished(__ -> b.setEffect(null));
            t.play();
        });
    }
}
