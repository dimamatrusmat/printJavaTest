package com.example.printtest;

import javafx.print.Printer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.ArrayList;
import java.util.Optional;

public class Print {

     static void wordToPDF(String input, String output, String dir) throws Exception {

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

    static void getAllPdfFile(){}

    static void printPDF(Optional<Printer> opt, String name, int duplexMod) {
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
