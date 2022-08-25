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

    requires pdfbox;
    requires java.desktop;
    requires itext;

    requires org.apache.poi.xwpf.converter.pdf;
    requires org.apache.poi.xwpf.converter.core;
    requires org.apache.poi.ooxml;

    opens com.example.printtest to javafx.fxml;
    exports com.example.printtest;
}