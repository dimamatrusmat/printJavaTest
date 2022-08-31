package com.example.printtest;

import javafx.print.Printer;
import org.apache.pdfbox.multipdf.Splitter;
import org.apache.pdfbox.pdmodel.PDDocument;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.Iterator;
import java.util.List;
import java.util.Optional;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import org.apache.pdfbox.printing.PDFPageable;


public class Print {

    final static String [] nameFile = new String[] {
            "Дело", // 0
            "Экзаменационный лист с фото", // 1
            "Заявление на поступление", // 2
            "Согласие на зачисление", // 3
            "Согласие на обработку персональных данных", // 4
            "Дог" // 5
    };

    static void getAllPdfFile() {
    }

    static void createMag(String name, String path) throws IOException {
        File file = new File(name);
        PDDocument document = PDDocument.load(file);
        splitPdf(document, getName(path, nameFile[0]), new int [] {1, 2});
        splitPdf(document, getName(path, nameFile[1]), new int [] {3, 4});
        splitPdf(document, getName(path, nameFile[2]), new int [] {5, 6});
        splitPdf(document, getName(path, nameFile[3]), new int [] {7, 7});
        splitPdf(document, getName(path, nameFile[4]), new int [] {8, 9});
        splitPdf(document, getName(path, nameFile[5]), new int [] {10, document.getPages().getCount()});

        document.close();
    }

    static void createBak(String name, String path) throws IOException {
        File file = new File(name);
        PDDocument document = PDDocument.load(file);

        splitPdf(document, getName(path, nameFile[0]), new int [] {1, 2});
        splitPdf(document, getName(path, nameFile[1]), new int [] {3, 4});
        splitPdf(document, getName(path, nameFile[2]), new int [] {5, 6});
        splitPdf(document, getName(path, nameFile[3]), new int [] {7, 7});
        splitPdf(document, getName(path, nameFile[4]), new int [] {8, 9});
        splitPdf(document, getName(path, nameFile[5]), new int [] {10, document.getPages().getCount()});

        document.close();
    }

    static void createSpo(String name, String path) throws IOException {
        File file = new File(name);
        PDDocument document = PDDocument.load(file);
        splitPdf(document, getName(path, nameFile[0]), new int [] {1, 2});
        splitPdf(document, getName(path, nameFile[2]), new int [] {3, 4});
        splitPdf(document, getName(path, nameFile[3]), new int [] {5, 5});
        splitPdf(document, getName(path, nameFile[4]), new int [] {6, 7});
        splitPdf(document, getName(path, nameFile[5]), new int [] {8, document.getPages().getCount()});

        document.close();
    }

    static void createOplate(String name, String path) throws IOException {
        File file = new File(name);
        PDDocument document = PDDocument.load(file);
        splitPdf(document, getName(path, "Квитанция"), new int [] {1, 1});
        document.close();
    }


    static void printPDF(Optional<Printer> opt, String name, int duplexMod, Boolean start) {


        if (opt.isPresent() && start) {
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

    public static void wordToPDF(String docxFilePath, String pdfFile) {
        System.out.println(docxFilePath);
        System.out.println(pdfFile);
        File docxFile = new File(docxFilePath);

        if (docxFile.exists()) {
            if (!docxFile.isDirectory()) {
                ActiveXComponent app = null;

                try {
                    app = new ActiveXComponent("Word.Application");
                    Dispatch documents = app.getProperty("Documents").toDispatch();
                    Dispatch document = Dispatch.call(documents, "Open", docxFilePath, false, true).toDispatch();
                    File target = new File(pdfFile);
                    if (target.exists()) {
                        target.delete();
                    }
                    Dispatch.call(document, "SaveAs", pdfFile, 17);
                    Variant f = new Variant(false);
                    Dispatch.call(document, "Close", f);

                } catch (Exception e) {
                    throw new RuntimeException("pdf convert failed.");
                } finally {
                    assert app != null;
                    app.invoke("Quit", new Variant[] {});
                }
            }
        }
    }

    public static void splitPdf(PDDocument document, String outfile,  int arr []) throws IOException {

        PDDocument pd = new PDDocument();

        if (arr[0] == arr[1]) {
            pd.addPage(document.getPage(arr[0]-1));
        } else {
            for (int i = arr[0] - 1; i <= arr[1] - 1; i ++) {
                pd.addPage(document.getPage(i));
            }

        }

        pd.save(outfile);
        pd.close();
    }


    static String getName(String path, String name) {
        return path + '/' + name + ".pdf";
    }

}
