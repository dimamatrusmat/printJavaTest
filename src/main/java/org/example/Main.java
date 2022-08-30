package org.example;


import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

import java.io.File;


public class Main {
    public static void main(String[] args) {

        String docxFile = "D:\\practing\\java\\convertor\\StartConver\\Договор №17.06.2022 Екатерина.docx";

        Main.convertDocx2pdf(docxFile);
    }


    public static void convertDocx2pdf(String docxFilePath){
            File docxFile = new File(docxFilePath);
            String pdfFile = docxFilePath.substring(0, docxFilePath.lastIndexOf(".docx")) + ".pdf";

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
}

