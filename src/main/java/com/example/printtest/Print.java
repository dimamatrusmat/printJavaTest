package com.example.printtest;

import javafx.print.Printer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.IOException;
import java.util.Optional;


public class Print {

    static void wordToPDF(String input, String output, String dir) throws Exception {

        String docPath = input;
        String pdfPath = dir + "/pdf/" + output;

//        try
//        {
//            // Load the document
//            WordDocument wd = new WordDocument(docPath, com.example.printtest.getWordConvertOptions());
//
//            // Save the document as a PDF file
//            wd.saveAsPDF(pdfPath);
//        }
//        catch (Throwable t)
//        {
//            t.printStackTrace();
//        }
//        try {
//            InputStream is = new FileInputStream(new File(docPath));
//            WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(is);
//            List sections = wordMLPackage.getDocumentModel().getSections();
//            for (int i = 0; i < sections.size(); i++) {
//                wordMLPackage.getDocumentModel().getSections().get(i).getPageDimensions();
//            }
//            Mapper fontMapper = new IdentityPlusMapper();
//            wordMLPackage.setFontMapper(fontMapper);
//            Docx4J.toPDF(wordMLPackage, new FileOutputStream(pdfPath));
//            System.err.println("Your Word Document Converted to PDF Successfully!");
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        InputStream doc = new FileInputStream(docPath);
//        ByteArrayOutputStream baos = new ByteArrayOutputStream();
//
//        XWPFDocument document = new XWPFDocument(doc);
//        PdfOptions options = PdfOptions.create();
//        PdfConverter.getInstance().convert(document, baos, options);
//        String base64_encoded = Base64.encodeBytes(baos.toByteArray());
//        System.out.println(base64_encoded);
//        String inputFile = docPath;
//        String outputFile = pdfPath;

//        System.out.println("inputFile:" + inputFile + ",outputFile:" + outputFile);
//        FileInputStream in = new FileInputStream(inputFile);
//
//        XWPFDocument document = new XWPFDocument(in);
//
//        System.out.println(document.getDocument());
//
//        File outFile = new File(outputFile);
//        OutputStream out = new FileOutputStream(outFile);
//        PdfOptions options = null;
//        PdfConverter.getInstance().convert(document, out, options);

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
