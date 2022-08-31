package com.example.printtest;

import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfWriter;
import javafx.print.Printer;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.Sides;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.*;
import java.util.List;
import java.util.Optional;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;


public class Print {

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

    public static void wordToPDF(String docxFilePath, String pdfFile){
        System.out.println(docxFilePath);
        System.out.println(pdfFile);
        File docxFile = new File(docxFilePath);

        if (docxFile.exists()) {
            if (!docxFile.isDirectory()) {
                ActiveXComponent app = null;

                long start = System.currentTimeMillis();
                try {
                    ComThread.InitMTA(true);
                    app = new ActiveXComponent("Word.Application");
                    Dispatch documents = app.getProperty("Documents").toDispatch();
                    Dispatch document = Dispatch.call(documents, "Open", docxFilePath, false, true).toDispatch();
                    File target = new File(pdfFile);
                    System.out.println(start);
                    if (target.exists()) {
                        target.delete();
                    }
                    Dispatch.call(document, "SaveAs", pdfFile, 17);
                    Dispatch.call(document, "Close", false);
                    long end = System.currentTimeMillis();
                    System.out.println(end);
                } catch (Exception e) {
                    throw new RuntimeException("pdf convert failed.");
                } finally {
                    if (app != null) {
                        app.invoke("Quit", new Variant[] {});
                    }
                    ComThread.Release();
                }
            }
        }
    }
//
//    public static void writePdfFile(String output, String file) throws FileNotFoundException, DocumentException {
//        FileOutputStream fileout = new FileOutputStream(file);
//        Document document = new Document();
//        PdfWriter.getInstance(document, fileout);
//        document.open();
//        String[] splitter = output.split("\\n");
//        for (int i = 0; i < splitter.length; i++) {
//            Chunk chunk = new Chunk(splitter[i]);
//            Font font = new Font();
//            font.setStyle(Font.UNDERLINE);
//            font.setStyle(Font.ITALIC);
//            chunk.setFont(font);
//            document.add(chunk);
//            Paragraph paragraph = new Paragraph();
//            paragraph.add("");
//            document.add(paragraph);
//        }
//        document.close();
//
//    }
//
//    public static String readDocxFile(String fileName) {
//        String output = "";
//        try {
//            File file = new File(fileName);
//            FileInputStream fis = new FileInputStream(file.getAbsolutePath());
//            XWPFDocument document = new XWPFDocument(fis);
//            List<XWPFParagraph> paragraphs = document.getParagraphs();
//            for (XWPFParagraph para : paragraphs) {
//                output = output + "\n" + para.getText() + "\n";
//            }
//            fis.close();
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return output;
//    }


}
