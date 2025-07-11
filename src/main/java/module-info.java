module ru.citlab24.protokol {
    requires javafx.controls;
    requires javafx.fxml;

    requires org.controlsfx.controls;
    requires java.desktop;
    requires com.formdev.flatlaf;
    requires org.kordamp.ikonli.swing;
    requires org.kordamp.ikonli.fontawesome5;

    opens ru.citlab24.protokol to javafx.fxml;
    exports ru.citlab24.protokol;
}