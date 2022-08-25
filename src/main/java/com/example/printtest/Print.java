package com.example.printtest;

import javafx.print.Printer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;

import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.OutputStream;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.Optional;

public class Print {

    static void wordToPDF(String input, String output, String dir) throws Exception {

        String docPath = input;
        String pdfPath = dir + "/pdf/" + output;

        String inputFile = docPath;
        String outputFile = pdfPath;

        System.out.println("inputFile:" + inputFile + ",outputFile:" + outputFile);
        FileInputStream in = new FileInputStream(inputFile);
        XWPFDocument document = new XWPFDocument(in);
        File outFile = new File(outputFile);
        OutputStream out = new FileOutputStream(outFile);
        PdfOptions options = null;
        PdfConverter.getInstance().convert(document, out, options);

//         POIFSFileSystem fs = null;
//         Document document = new Document();
//
//         try {
//             System.out.println("Starting the test");
//             fs = new POIFSFileSystem(new FileInputStream(docPath));
//
//             HWPFDocument doc = new HWPFDocument(fs);
//             WordExtractor we = new WordExtractor(doc);
//
//             OutputStream file = new FileOutputStream(new File(pdfPath));
//
//             PdfWriter writer = PdfWriter.getInstance(document, file);
//
//             org.apache.poi.hwpf.usermodel.Range range = doc.getRange();
//             document.open();
//             writer.setPageEmpty(true);
//             document.newPage();
//             writer.setPageEmpty(true);
//
//             String[] paragraphs = we.getParagraphText();
//             for (int i = 0; i < paragraphs.length; i++) {
//
//                 org.apache.poi.hwpf.usermodel.Paragraph pr = range.getParagraph(i);
//                 // CharacterRun run = pr.getCharacterRun(i);
//                 // run.setBold(true);
//                 // run.setCapitalized(true);
//                 // run.setItalic(true);
//                 paragraphs[i] = paragraphs[i].replaceAll("\\cM?\r?\n", "");
//                 System.out.println("Length:" + paragraphs[i].length());
//                 System.out.println("Paragraph" + i + ": " + paragraphs[i].toString());
//
//                 document.add(new Paragraph(paragraphs[i]));
//             }
//
//             System.out.println("Document testing completed");
//         } catch (Exception e) {
//             System.out.println("Exception during test");
//             e.printStackTrace();
//         } finally {
//             // close the document
//             document.close();
//         }



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
