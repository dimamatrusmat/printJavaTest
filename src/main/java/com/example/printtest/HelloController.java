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
        Boolean fileChose = true;

        pringBtn.setOnAction(event -> {
            File outdir = new File("output/");
            System.out.println(outdir.getAbsolutePath());
            ArrayList <String> arrout = listFolderForFolder(outdir);

            ChoiceDialog dialogMake = new ChoiceDialog(arrout.get(0), arrout);
            dialogMake.setHeaderText("Выбери кого распечатать!");
            dialogMake.setContentText("Выберите дело из списка");
            dialogMake.setTitle("Выбор дела");
            String getDo = dialogMake.showAndWait().get().toString();
            File dir = new File(outdir.getAbsoluteFile().toString() + '/' + getDo);

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

            ArrayList <String> arrpdf = listFolderForFolder(new File(path.toString()));

            ChoiceDialog dialogPrint = new ChoiceDialog("Всё дело", arrpdf);
            dialogPrint.setHeaderText("Выбери файл для печати!");
            dialogPrint.setContentText("Выберите файлы из списка");
            dialogPrint.setTitle("Выбор файла");

            ArrayList <String> filePrint = new ArrayList<>();
            // Выбор что печатать
            if (dialogPrint.showAndWait().get().toString() == "Всё дело") {
                arrpdf.forEach((pdf) -> {filePrint.add(pdf);});
            } else {
                filePrint.add(dialogPrint.showAndWait().get().toString());
            }

//             Выбираем принтер
            ChoiceDialog dialog = new ChoiceDialog(Printer.getDefaultPrinter(), Printer.getAllPrinters());
            dialog.setHeaderText("Выбери принтер!");
            dialog.setContentText("Выберите принтер из списка");
            dialog.setTitle("Выбор принтера");
            Optional<Printer> opt = dialog.showAndWait();

            // Расбираем файлы
            filePrint.forEach((file) -> {
                file = path.toString()+'/'+ file;

                if (file.contains("МАГ")) {
                    try {
                        Print.createMag(file, path.toString());
                        new File (file).delete();

                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }

                } else if (file.contains("БАК")) {
                    try {
                        Print.createBak(file, path.toString());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                    new File (file).delete();

                } else if (file.contains("СПО")) {
                    try {
                        Print.createSpo(file, path.toString());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                    new File (file).delete();

                } else if (file.contains("Квитанция")) {
                    try {
                        Print.createOplate(file, path.toString());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }

                }


            });


            ArrayList <String> arrprintpdf = listFolderForFolder(new File(path.toString()));

            arrprintpdf.forEach((file) -> {
                int duplexMod = 1;
                Boolean start = true;

                if (file.contains("фото")){
                    duplexMod = 2;
                } else if (file.contains("Договор")) {
                    start = false;
                }
                file = path.toString() +'/'+ file;
                Print.printPDF(opt, file, duplexMod, start);
            }
            );

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

    public ArrayList<String> listFolderForFolder(final File folder) {

        ArrayList <String> arr = new ArrayList();

        for (final File fileEntry : folder.listFiles()) {
            String name = fileEntry.getName();
            arr.add(name);
        }

        return arr;
    }
}