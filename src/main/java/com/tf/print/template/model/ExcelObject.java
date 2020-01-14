package com.tf.print.template.model;

import cn.hutool.poi.excel.ExcelPicUtil;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.property.TextAlignment;
import com.tf.print.template.excel.ExcelExReader;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.element.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 参照项目https://github.com/caryyu/excel2pdf 进行了修改
 * @ClassName ExcelObject
 * @Description TODO
 * @Author kyjonny
 * @Date 6/1/2020 10:51 上午
 **/
public class ExcelObject extends PdfTool{

    private ExcelExReader excelReader;

    private PageSize pageSize = PageSize.A4;

    //中文字体
    static PdfFont bfChinese = null;
    static {

        try {
            bfChinese = PdfFontFactory.createFont("STSong-Light", "UniGB-UCS2-H",true);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public ExcelObject(ExcelExReader excelReader, OutputStream os) {
        this.excelReader = excelReader;
        this.os = os;

    }

    public void setPageSize(PageSize pageSize) {
        this.pageSize = pageSize;
    }

    public void convert() throws IOException {
        getDocument().getPdfDocument().setDefaultPageSize(this.pageSize);
        getDocument().setMargins(5,5,5,5);
        Table table = this.toCreatePdfTableInfo();
        getDocument().add(table);
        getDocument().close();
    }

    /**
     * 创建pdf表格信息
     * @return
     */
    private Table toCreatePdfTableInfo() {

        Table table = new Table(getPdfPTableWidths()).useAllAvailableWidth().setFixedLayout();
        convertContent(table);
        return table;
    }

    private void convertContent(Table table) {

        int rows = excelReader.getSheet().getPhysicalNumberOfRows();
        List<Cell> cells = new ArrayList<>();
        float[] widths = null;
//        float mw = 0;
        Map<String, MergeCelInfo> skipCells = new HashMap<>();
        for (int i = 0; i < rows; i++) {
            Row row = excelReader.getSheet().getRow(i);
            int columns = row.getLastCellNum();

//            float[] cws = new float[columns];
            for (int j = 0; j < columns; j++) {

                MergeCelInfo mergeCelInfo = skipCells.get(getKey(i,j));
                if(mergeCelInfo !=null){
                    if(mergeCelInfo.isSkipRow()){
                        break;
                    }
                }

                org.apache.poi.ss.usermodel.Cell cell = row.getCell(j);
                if (cell == null){
                    cell = row.createCell(j);
                }

                float cw = getColumnWidth(cell);
//                cws[cell.getColumnIndex()] = cw;

                cell.setCellType(CellType.STRING);

                int rowspan = 1;
                int colspan = 1;
                if(mergeCelInfo == null){
                    CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
                    if (range != null) {
                        rowspan = range.getLastRow() - range.getFirstRow() + 1;
                        colspan = range.getLastColumn() - range.getFirstColumn() + 1;
                        for(int num=range.getFirstRow();num <= range.getLastRow();num++){
                            MergeCelInfo mergeCelInfoTmp = new MergeCelInfo();
                            mergeCelInfoTmp.setRow(num)
                                    .setCol(range.getFirstColumn())
                                    .setStep(colspan)
                                    .setSkipRow(columns == colspan);
                            skipCells.put(getKey(num,range.getFirstColumn()),mergeCelInfoTmp);
                        }
                    }
                }else{
                    colspan = mergeCelInfo.getStep();
                    j += colspan - 1;
                    continue;
                }

                Cell pdfpCell = new Cell(rowspan,colspan);
                pdfpCell.setPaddingTop(0);
                pdfpCell.setPaddingBottom(0);
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setTextAlignment(getTextAlignment(cell.getCellStyle().getAlignment()));

                if (excelReader.getSheet().getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
                    pdfpCell.setHeight(this.getPixelHeight(row.getHeightInPoints()));
                }
                addImageByPOICell(pdfpCell , cell , cw);
                pdfpCell.add(getValue(cell,null,10));
                addBorderByExcel(pdfpCell, cell.getCellStyle(),i,j);

                cells.add(pdfpCell);
                j += colspan - 1;
            }
//
//            float rw = 0;
//            for (int j = 0; j < cws.length; j++) {
//                rw += cws[j];
//            }
//            if (rw > mw ||  mw == 0) {
//                widths = cws;
//                mw = rw;
//            }
        }

        for (Cell pdfpCell : cells) {
            table.addCell(pdfpCell);
        }
    }

    /**
     * 设置文本对齐方式
     * @param alignment
     * @return
     */
    private TextAlignment getTextAlignment(HorizontalAlignment alignment) {
        TextAlignment result = TextAlignment.LEFT;
        if (alignment == HorizontalAlignment.LEFT) {
            result = TextAlignment.LEFT;
        }
        if (alignment == HorizontalAlignment.RIGHT) {
            result = TextAlignment.RIGHT;
        }
        if (alignment == HorizontalAlignment.DISTRIBUTED) {
            result = TextAlignment.JUSTIFIED_ALL;
        }
        if (alignment == HorizontalAlignment.CENTER) {
            result = TextAlignment.CENTER;
        }
        return result;
    }

    /**
     * 根据行列信息获取key
     * @param row
     * @param col
     * @return
     */
    private String getKey(int row,int col){
        return String.format("%d-%d",row,col);
    }

    private Paragraph getValue(org.apache.poi.ss.usermodel.Cell cell, String language, float size) {

        PdfFont f = getFontForThisLanguage(language);
        Paragraph paragraph = new Paragraph(cell.getStringCellValue()).setFont(f);
        paragraph.setFontSize(size);
        return paragraph;
    }

    private PdfFont getFontForThisLanguage(String language) {
        if (language == null) {
            return bfChinese;
        }
        switch (language) {
            default: {
                return bfChinese;
            }
        }
    }


    protected void addImageByPOICell(Cell pdfpCell , org.apache.poi.ss.usermodel.Cell cell , float cellWidth){
        byte[] bytes = getCellImage(cell);
        if(bytes != null){
            pdfpCell.setVerticalAlignment(com.itextpdf.layout.property.VerticalAlignment.MIDDLE);
            pdfpCell.setHorizontalAlignment(com.itextpdf.layout.property.HorizontalAlignment.CENTER);
            Image image = new Image(ImageDataFactory.create(bytes));
            pdfpCell.add(image.setAutoScale(true));
        }
    }

    private byte[] getCellImage(org.apache.poi.ss.usermodel.Cell cell) {
        Map<String, PictureData> picMap = ExcelPicUtil.getPicMap(excelReader.getWorkbook(),excelReader.getWorkbook().getSheetIndex(cell.getSheet().getSheetName()));
        PictureData picData = picMap.get(String.format("%s_%s",cell.getRowIndex(),cell.getColumnIndex()));
        if(picData == null){
            return null;
        }
        return picData.getData();
    }


    /**
     * 获取Excel的列宽像素
     * @param cell
     * @return
     */
    protected float getColumnWidth(org.apache.poi.ss.usermodel.Cell cell) {
//        int poiCWidth = excelReader.getSheet().getColumnWidth(cell.getColumnIndex());
//        int colWidthpoi = poiCWidth;
//        int widthPixel = 0;
//        if (colWidthpoi >= 416) {
//            widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
//        } else {
//            widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
//        }
//        return widthPixel;
        return excelReader.getSheet().getColumnWidthInPixels(cell.getColumnIndex());
    }

    protected void addBorderByExcel(Cell cell , CellStyle style,int row,int col) {
        cell.setBorderTop(style.getBorderTop().getCode() ==0? Border.NO_BORDER :new SolidBorder(style.getBorderTop().getCode()));
        cell.setBorderLeft(style.getBorderLeft().getCode() ==0? Border.NO_BORDER:new SolidBorder(style.getBorderLeft().getCode()));
        if(cell.getColspan()>1){
            style = excelReader.getCell(col+cell.getColspan()-1,row).getCellStyle();
            cell.setBorderRight(style.getBorderRight().getCode() == 0? Border.NO_BORDER:new SolidBorder(style.getBorderRight().getCode()));
        }else{
            cell.setBorderRight(style.getBorderRight().getCode() == 0? Border.NO_BORDER:new SolidBorder(style.getBorderRight().getCode()));
        }
        if(cell.getRowspan() >1){
            style = excelReader.getCell(col,row+cell.getRowspan()-1).getCellStyle();
            cell.setBorderBottom(style.getBorderBottom().getCode() ==0? Border.NO_BORDER:new SolidBorder(style.getBorderBottom().getCode()));
        }else{
            cell.setBorderBottom(style.getBorderBottom().getCode() ==0? Border.NO_BORDER:new SolidBorder(style.getBorderBottom().getCode()));
        }
    }


    protected float getPixelHeight(float poiHeight){
        float pixel = poiHeight / 28.6f * 26f;
        return pixel;
    }

    protected CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex) {
        CellRangeAddress result = null;
        Sheet sheet = excelReader.getSheet();
        int num = sheet.getNumMergedRegions();
        for (int i = 0; i < num; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
                result = range;
                break;
            }
        }
        return result;
    }


    /**
     * 获取表格宽度
     * @return
     */
    private float[] getPdfPTableWidths(){
        float[] widths= new float[excelReader.getColumnCount()];
        for(int i=0 ;i<excelReader.getColumnCount();i++){
            widths[i] = getDocument().getPdfDocument().getDefaultPageSize().getWidth()/excelReader.getCountWidthInPixels();
        }
        return widths;
    }

    protected com.itextpdf.layout.property.VerticalAlignment getVAlignByExcel(VerticalAlignment align) {
        com.itextpdf.layout.property.VerticalAlignment result = com.itextpdf.layout.property.VerticalAlignment.MIDDLE;
        if (align == VerticalAlignment.BOTTOM) {
            result = com.itextpdf.layout.property.VerticalAlignment.BOTTOM;
        }
        if (align == VerticalAlignment.CENTER) {
            result = com.itextpdf.layout.property.VerticalAlignment.MIDDLE;
        }
        if (align == VerticalAlignment.TOP) {
            result = com.itextpdf.layout.property.VerticalAlignment.TOP;
        }
        return result;
    }

    protected com.itextpdf.layout.property.HorizontalAlignment getHAlignByExcel(HorizontalAlignment align) {
        com.itextpdf.layout.property.HorizontalAlignment result= com.itextpdf.layout.property.HorizontalAlignment.CENTER;
        if (align == HorizontalAlignment.LEFT) {
            result = com.itextpdf.layout.property.HorizontalAlignment.LEFT;
        }
        if (align == HorizontalAlignment.RIGHT) {
            result = com.itextpdf.layout.property.HorizontalAlignment.RIGHT;
        }
        if (align == HorizontalAlignment.CENTER) {
            result = com.itextpdf.layout.property.HorizontalAlignment.CENTER;
        }
        return result;
    }
}
