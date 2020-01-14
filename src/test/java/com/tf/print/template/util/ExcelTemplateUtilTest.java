package com.tf.print.template.util;

import cn.hutool.core.lang.Console;
import com.itextpdf.kernel.geom.PageSize;
import com.tf.print.template.excel.ExcelExReader;
import com.tf.print.template.model.ExcelObject;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;


public class ExcelTemplateUtilTest {

    private String excelPath ="e:/asd.xls";
    private String excelPath2 ="e:/小票.xls";

    @org.junit.Before
    public void setUp() throws Exception {
    }

    @org.junit.After
    public void tearDown() throws Exception {
    }

    @org.junit.Test
    public void getReader() {

        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        System.out.println(templateReader.isXlsx());
        int firstRow = templateReader.getSheet().getFirstRowNum();
        int lastRow = templateReader.getSheet().getLastRowNum();
        for(int i=firstRow;i<=lastRow;i++){
            int firstCell = templateReader.getSheet().getRow(i).getFirstCellNum();
            int lastCell = templateReader.getSheet().getRow(i).getLastCellNum();
            for(int j=firstCell;j<=lastCell;j++){
                Cell cell = templateReader.getSheet().getRow(i).getCell(j);
                System.out.println(cell);
            }
            Console.log("cell {}<{}",firstCell,lastCell);
        }
        Console.log("{},{}",firstRow,lastRow);
    }


    @Test
    public void getMergeCell(){
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        List<CellRangeAddress> mergeList =templateReader.getSheet().getMergedRegions();
        mergeList.forEach(item ->{
            Console.log("合并单元格信息，开始行{},结束行{}.开始列{}，结束列{}, 单元格内容{}",item.getFirstRow(),item.getLastRow(),item.getFirstColumn(),item.getLastColumn(),item.formatAsString());
        });
    }

    @Test
    public void getPicture(){
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        List<HSSFPictureData> pictures = ((HSSFSheet) templateReader.getSheet()).getWorkbook().getAllPictures();
        if(pictures.size()>0){
            for (HSSFShape shape : ((HSSFSheet) templateReader.getSheet()).getDrawingPatriarch().getChildren()){
                HSSFClientAnchor hssfClientAnchor = (HSSFClientAnchor) shape.getAnchor();

                if (shape instanceof HSSFPicture) {
                    HSSFPicture pic = (HSSFPicture) shape;
                    int pictureIndex = pic.getPictureIndex() - 1;
                    HSSFPictureData picData = pictures.get(pictureIndex);
//                    String picIndex = String.valueOf(sheetNum) + "_"
//                            + String.valueOf(anchor.getRow1()) + "_"
//                            + String.valueOf(anchor.getCol1());
//                    sheetIndexPicMap.put(picIndex, picData);
                    System.out.println(picData.getData());
                    System.out.println(picData.getPictureType());
                }
            }
        }
    }

    @Test
    public void convertImage() throws IOException {
//        Rectangle rectangle = new Rectangle(0,0,(float) (74*72/25.4),(float) (130*72/25.4));

//        17.6     16
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath);
        ExcelObject excelObject = new ExcelObject(templateReader,new FileOutputStream(new File("e:/templateConvertPdf.pdf")));
//        excelObject.setAutoFit(false);
//        excelObject.setPageSize(new PageSize((float) (8*2.54*72), (float) (12*2.54*72)));
//        excelObject.setPageSize(rectangle);
        excelObject.setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));

        excelObject.convert();
    }

    @Test
    public void convertImage2() throws IOException {
//        Rectangle rectangle = new Rectangle(0,0,(float) (74*72/25.4),(float) (130*72/25.4));

//        17.6     16
        ExcelExReader templateReader = ExcelTemplateUtil.getReader(excelPath2,0);
        ExcelObject excelObject = new ExcelObject(templateReader,new FileOutputStream(new File("e:/templateConvertPdf2.pdf")));
//        excelObject.setAutoFit(false);
//        excelObject.setPageSize(new PageSize((float) (8*2.54*72), (float) (12*2.54*72)));
//        excelObject.setPageSize(rectangle);
        excelObject.setPageSize(new PageSize((float) (74*72/25.4),(float) (130*72/25.4)));

        excelObject.convert();
    }
}
