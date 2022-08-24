module com.example.printtest {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires validatorfx;
    requires org.kordamp.ikonli.javafx;
    requires org.kordamp.bootstrapfx.core;
    requires eu.hansolo.tilesfx;
    requires com.almasb.fxgl.all;
    requires pdfbox;
    requires java.desktop;
    requires com.aspose.words;


    opens com.example.printtest to javafx.fxml;
    exports com.example.printtest;
}