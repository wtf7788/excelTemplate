package com.tf.print.template.util;

import com.itextpdf.kernel.geom.PageSize;
import com.tf.print.template.excel.ExcelExReader;
import com.tf.print.template.model.ExcelObject;
import org.junit.Test;
import sun.misc.BASE64Decoder;

import java.io.*;


public class ExcelTemplateUtilTest {

    private String excelPath ="/Users/kyjonny/Desktop/tmp/小票.xls";
    private String excelPath2 ="/Users/kyjonny/Desktop/tmp/小票.xls";

    @org.junit.Before
    public void setUp() throws Exception {
    }

    @org.junit.After
    public void tearDown() throws Exception {
    }


    @Test
    public void getPdf() throws IOException {
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        ExcelObject excelObject = new ExcelObject(templateReader,"/Users/kyjonny/Desktop/tmp/ll.pdf");
        excelObject.setDpi(186)
            .setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));
        excelObject.convertPdf();
    }

    @Test
    public void getPdf2() throws IOException {
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        ExcelObject excelObject = new ExcelObject(templateReader);
        excelObject.setDpi(186)
                .setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));
        excelObject.convertPdf();

        System.out.println(excelObject.getBase64());
        decodeBase64ToFile(excelObject.getBase64(),"/Users/kyjonny/Desktop/tmp/","base64.pdf");

    }

    public static void decodeBase64ToFile(String base64, String path,String fileName) {
        BASE64Decoder decoder = new BASE64Decoder();
        try {
            FileOutputStream write = new FileOutputStream(new File(path + fileName));
            byte[] decoderBytes = decoder.decodeBuffer(base64);
            write.write(decoderBytes);
            write.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void getImage() throws IOException {
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        ExcelObject excelObject = new ExcelObject(templateReader,"/Users/kyjonny/Desktop/tmp/test.png");
        excelObject.setDpi(186)
                .setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));
        excelObject.convertImg();
    }

    @Test
    public void getImage2() throws IOException {
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        ExcelObject excelObject = new ExcelObject(templateReader);
        excelObject.setDpi(186)
                .setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));
        excelObject.convertImg();
        System.out.println(excelObject.getBase64());
        decodeBase64ToFile(excelObject.getBase64(),"/Users/kyjonny/Desktop/tmp/","base64---2.png");

    }

}
