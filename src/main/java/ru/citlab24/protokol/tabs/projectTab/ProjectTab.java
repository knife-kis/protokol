package ru.citlab24.protokol.tabs.projectTab;

import javafx.application.Platform;
import javafx.embed.swing.JFXPanel;

import javax.swing.*;
import java.awt.*;

public class ProjectTab extends JPanel {
    private final Runnable onLoad;
    private final Runnable onSave;

    public ProjectTab(Runnable onLoad, Runnable onSave) {
        this.onLoad = onLoad;
        this.onSave = onSave;
        initComponents();
    }

    private void initComponents() {
        setLayout(new BorderLayout());
        add(createActionButtons(), BorderLayout.NORTH);
    }

    private JPanel createActionButtons() {
        JPanel wrap = new JPanel(new BorderLayout());
        final JFXPanel fx = new JFXPanel();
        wrap.add(fx, BorderLayout.CENTER);

        Platform.runLater(() -> {
            javafx.scene.control.Button btnLoad = new javafx.scene.control.Button("Загрузить проект");
            javafx.scene.control.Button btnSave = new javafx.scene.control.Button("Сохранить проект");

            org.kordamp.ikonli.javafx.FontIcon icLoad = new org.kordamp.ikonli.javafx.FontIcon("fas-folder-open");
            org.kordamp.ikonli.javafx.FontIcon icSave = new org.kordamp.ikonli.javafx.FontIcon("fas-save");
            icLoad.setIconSize(16);
            icSave.setIconSize(16);
            btnLoad.setGraphic(icLoad);
            btnSave.setGraphic(icSave);

            btnLoad.getStyleClass().addAll("button", "btn-load");
            btnSave.getStyleClass().addAll("button", "btn-save");

            javafx.scene.layout.HBox box = new javafx.scene.layout.HBox(10, btnLoad, btnSave);
            box.getStyleClass().addAll("controls-bar", "theme-light");
            for (javafx.scene.control.Button b : java.util.List.of(btnLoad, btnSave)) {
                b.setMaxWidth(Double.MAX_VALUE);
                javafx.scene.layout.HBox.setHgrow(b, javafx.scene.layout.Priority.ALWAYS);
            }

            javafx.scene.Scene scene = new javafx.scene.Scene(box);
            scene.getStylesheets().add(
                    java.util.Objects.requireNonNull(getClass().getResource("/ui/protokol.css")).toExternalForm()
            );
            fx.setScene(scene);

            final String colLoad = "#3949ab";
            final String colSave = "#43a047";
            final java.util.List<javafx.scene.control.Button> all =
                    java.util.List.of(btnLoad, btnSave);

            final java.util.function.BiConsumer<javafx.scene.control.Button, String> markActive =
                    (btn, hex) -> {
                        all.forEach(ru.citlab24.protokol.ui.fx.ActiveGlow::clear);
                        ru.citlab24.protokol.ui.fx.ActiveGlow.setActive(btn, hex, 0.5);
                    };

            btnLoad.setOnAction(ev -> {
                markActive.accept(btnLoad, colLoad);
                Runnable action = (onLoad != null) ? onLoad : () -> {};
                javax.swing.SwingUtilities.invokeLater(action);
            });
            btnSave.setOnAction(ev -> {
                markActive.accept(btnSave, colSave);
                Runnable action = (onSave != null) ? onSave : () -> {};
                javax.swing.SwingUtilities.invokeLater(action);
            });
        });

        return wrap;
    }
}
