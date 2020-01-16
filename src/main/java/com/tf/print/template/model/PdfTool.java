package com.tf.print.template.model;

import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;

import java.io.ByteArrayOutputStream;

/**
 * @ClassName PdfTool
 * @Description TODO
 * @Author kyjonny
 * @Date 6/1/2020 10:55 上午
 **/
public class PdfTool {

    protected Document document;

    protected ByteArrayOutputStream os;

    public Document getDocument() {

        if (document == null) {
            PdfDocument pdf = new PdfDocument(new PdfWriter(os));
            document = new Document(pdf);
        }
        return document;
    }
}
