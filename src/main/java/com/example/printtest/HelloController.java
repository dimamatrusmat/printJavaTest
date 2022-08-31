package com.example.printtest;

import javafx.fxml.FXML;
import javafx.print.Printer;
import javafx.scene.control.ChoiceDialog;
import javafx.stage.DirectoryChooser;
import javafx.scene.control.Button;
import javafx.stage.FileChooser;


import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Optional;

public class HelloController {

    @FXML
    private Button pringBtn;

    @FXML
    void initialize() {
        System.out.println(1);
        int duplexMod = 1;
        Boolean fileChose = true;

        pringBtn.setOnAction(event -> {

            System.out.println(1);
            DirectoryChooser directoryChooser = new DirectoryChooser();

            File dir = directoryChooser.showDialog(pringBtn.getScene().getWindow());
            ArrayList <String> arr = listFilesForFolder(dir);

            Path path = Paths.get(dir+"/pdf");

            if (!Files.exists(path)) {
                new File(dir+"/pdf").mkdirs();
            }

            arr.forEach((word) -> {try {
                     Print.wordToPDF(
                    dir.getAbsolutePath()+"/"+word,
                    dir.getAbsolutePath()+"/pdf/"+word.substring(0, word.length() - 4) + "pdf"
                    );
                        } catch (Exception e) {
                            throw new RuntimeException(e);
                        }
                    }
            );

//            //Ищем pdf файл
//            if (fileChose == true) {
//
//            }
//
//            FileChooser fileChooser = new FileChooser();
//            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("PDF", "*.pdf","*.PDF");
//            fileChooser.getExtensionFilters().add(extFilter);
//            File files = fileChooser.showOpenDialog(pringBtn.getScene().getWindow());
//            String file = files.getAbsolutePath();
//            System.out.println(files.getAbsolutePath());
//
//            // Выбираем принтер
//            ChoiceDialog dialog = new ChoiceDialog(Printer.getDefaultPrinter(), Printer.getAllPrinters());
//            dialog.setHeaderText("Выбери принтер!");
//            dialog.setContentText("Выберите принтер из списка");
//            dialog.setTitle("Выбор принтера");
//            Optional<Printer> opt = dialog.showAndWait();
//
//
//            // Print file
//            printPDF(opt, file, duplexMod);
        });

    }
    public ArrayList<String> listFilesForFolder(final File folder) {

        ArrayList <String> arr = new ArrayList();

        for (final File fileEntry : folder.listFiles()) {
            if (fileEntry.isDirectory()) {
                listFilesForFolder(fileEntry);
            } else {
                String name = fileEntry.getName();

                if (name.contains(".docx")) {
                    System.out.println(name);
                    arr.add(name);
                }
            }
        }

        return arr;
    }
}