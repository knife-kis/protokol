module ru.citlab24.protokol {
    requires javafx.controls;
    requires javafx.fxml;

    requires org.controlsfx.controls;
    requires java.desktop;
    requires com.formdev.flatlaf;
    requires org.kordamp.ikonli.swing;
    requires org.kordamp.ikonli.fontawesome5;
    requires java.sql;
//    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;

    opens ru.citlab24.protokol to javafx.fxml;
    exports ru.citlab24.protokol;
}