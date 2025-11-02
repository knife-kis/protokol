module ru.citlab24.protokol {
    requires javafx.controls;
    requires javafx.graphics;
    requires javafx.fxml;
    requires javafx.swing;
    requires org.kordamp.ikonli.javafx;
    requires org.kordamp.ikonli.fontawesome5;

    requires org.controlsfx.controls;
    requires java.desktop;
    requires com.formdev.flatlaf;
    requires org.kordamp.ikonli.swing;
    requires java.sql;
//    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires org.slf4j;

    opens ru.citlab24.protokol to javafx.fxml;
    exports ru.citlab24.protokol;
}