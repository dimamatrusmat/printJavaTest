package com.example.printtest;

import javafx.fxml.FXML;
import javafx.print.Printer;
import javafx.stage.DirectoryChooser;


import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import javafx.scene.control.Button;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;

import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

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
                     wordToPDF(
                    dir.getAbsolutePath()+"/"+word,
                    word.substring(0, word.length() - 4) + "pdf",
                    dir.getAbsolutePath()
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

    private static void wordToPDF(String input, String output, String dir) throws Exception {
        String docPath = input;
        String pdfPath = dir+"/pdf/"+output;

        try {
            InputStream doc = new FileInputStream(new File(docPath));
            XWPFDocument document = new XWPFDocument(doc);
            PdfOptions options = PdfOptions.create();
            OutputStream out = new FileOutputStream(new File(pdfPath));
            PdfConverter.getInstance().convert(document, out, options);
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }


//        try {
//            InputStream templateInputStream = new FileInputStream(input);
//            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(templateInputStream);
//            MainDocumentPart documentPart = wordMLPackage.getMainDocumentPart();
//
//            FileOutputStream os = new FileOutputStream(dir+"/pdf/"+output);
//            Docx4J.toPDF(wordMLPackage,os);
//            os.flush();
//            os.close();
//        } catch (Throwable e) {
//
//            e.printStackTrace();
//        }
//

    }

    private static void getAllPdfFile(){}

    private static void printPDF(Optional<Printer> opt, String name, int duplexMod) {
        if (opt.isPresent()) {
            Printer printer = opt.get();
            PrinterJob job = PrinterJob.getPrinterJob();
            try {
                PDDocument document = PDDocument.load(new File(name));
                PrintRequestAttributeSet pras = new HashPrintRequestAttributeSet();
                if (duplexMod == 2) {
                    pras.add(Sides.TWO_SIDED_SHORT_EDGE);
                } else {
                    pras.add(Sides.TWO_SIDED_LONG_EDGE);
                }

                job.setPageable(new PDFPageable(document));
                job.setPrintService(findPrintService(printer.getName()));
                job.print(pras);
            } catch (IOException | PrinterException e) {
                throw new RuntimeException(e);
            }
        }
    }

    private static PrintService findPrintService(String printerName) {
        PrintService[] printServices = PrintServiceLookup.lookupPrintServices(null, null);
        for (PrintService printService : printServices) {
            if (printService.getName().trim().equals(printerName)) {
                return printService;
            }
        }
        return null;
    }

}