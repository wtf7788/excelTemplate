package com.tf.print.template.model;

import cn.hutool.core.util.ReflectUtil;
import cn.hutool.core.util.StrUtil;
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
import com.tf.print.template.exception.TemplateException;
import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.pdfbox.tools.imageio.ImageIOUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import sun.misc.BASE64Encoder;

import java.awt.image.BufferedImage;
import java.io.*;
import java.util.*;

/**
 * 参照项目https://github.com/caryyu/excel2pdf 进行了修改
 * @ClassName ExcelObject
 * @Description TODO
 * @Author kyjonny
 * @Date 6/1/2020 10:51 上午
 **/
@Slf4j
public class ExcelObject extends PdfTool{

    private ExcelExReader excelReader;

    private PageSize pageSize = PageSize.A4;

    private String DELIMITER_PLACEHOLDER_START = "${";

    private String DELIMITER_PLACEHOLDER_END = "}";

    private Object data;

    /**
     * 转换图片时默认dpi值为300
     */
    private int dpi = 300;

    /**
     * 生产图片时，渲染的图像类型
     */
    private ImageType colourType = ImageType.RGB;

    /**
     * 导出图片类型
     */
    private PictureType pictureType = PictureType.PNG;

    private String filePath;

    private ByteArrayOutputStream byteArrayOutputStream;

    //中文字体
    static PdfFont bfChinese = null;

    static {
        try {
            bfChinese = PdfFontFactory.createFont("STSong-Light", "UniGB-UCS2-H",true);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public ExcelObject(ExcelExReader excelReader) {
        this.excelReader = excelReader;
        this.os = new ByteArrayOutputStream();
    }

    public ExcelObject(ExcelExReader excelReader, String path) throws IOException {
        this.excelReader = excelReader;
        this.os = new ByteArrayOutputStream();
        this.filePath = path;
    }

    public String getBase64(){
        BASE64Encoder encoder = new BASE64Encoder();
        return encoder.encode(this.byteArrayOutputStream.toByteArray());
    }

    public void convertImg() throws IOException {
        convert();
        if(StrUtil.isNotBlank(filePath)){
            String formatName = filePath.substring(filePath.lastIndexOf(46) + 1);
            if(StrUtil.isNotBlank(formatName)){
                try {
                    this.pictureType = PictureType.valueOf(formatName);
                } catch (IllegalArgumentException e) {
                    throw new TemplateException("不支持当前类型的图片转化");
                }
            }
        }
        pdfConvertImage();
        if(StrUtil.isNotBlank(filePath) ){
            this.byteArrayOutputStream.writeTo(new FileOutputStream(filePath));
        }
    }

    private void pdfConvertImage() {
        ByteArrayOutputStream outStream = new ByteArrayOutputStream();
        outStream = (ByteArrayOutputStream) this.os;
        try (final PDDocument document = PDDocument.load( new ByteArrayInputStream(outStream.toByteArray()))){
            PDFRenderer pdfRenderer = new PDFRenderer(document);
            for (int page = 0; page < document.getNumberOfPages(); ++page){
                BufferedImage bim = pdfRenderer.renderImageWithDPI(page, dpi, colourType);
                float compressionQuality = 1.0F;
                if (this.pictureType.equals(PictureType.png) || this.pictureType.equals(PictureType.PNG)) {
                    compressionQuality = 0.0F;
                }
                try {
                    this.byteArrayOutputStream = new ByteArrayOutputStream();
                    ImageIOUtil.writeImage(bim, this.pictureType.name(), this.byteArrayOutputStream, 186, compressionQuality);
                } finally {
                    this.byteArrayOutputStream.close();
                }
            }
            document.close();
        } catch (IOException e){
            log.error("Exception while trying to create pdf document - {}" ,e);
        }
    }

    public ExcelObject setPageSize(PageSize pageSize) {
        this.pageSize = pageSize;
        return this;
    }

    public void convertPdf() throws IOException {
        convert();
        this.byteArrayOutputStream = (ByteArrayOutputStream) this.os;
        if(StrUtil.isNotBlank(filePath)){
            this.byteArrayOutputStream.writeTo(new FileOutputStream(filePath));
        }
    }


    private void convert() throws IOException {
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
        Map<String, MergeCelInfo> skipCells = new HashMap<>();
        for (int i = 0; i < rows; i++) {
            Row row = excelReader.getSheet().getRow(i);
            int columns = row.getLastCellNum();

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
                pdfpCell.setPaddingTop(0f);
                pdfpCell.setPaddingBottom(0f);
                pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
                pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
                pdfpCell.setTextAlignment(getTextAlignment(cell.getCellStyle().getAlignment()));

                if (excelReader.getSheet().getDefaultRowHeightInPoints() != row.getHeightInPoints()) {
                    pdfpCell.setHeight(this.getPixelHeight(row.getHeightInPoints()));
                }
                addImageByPOICell(pdfpCell , cell , cw);
                pdfpCell.add(getValue(cell,null,11));
                addBorderByExcel(pdfpCell, cell.getCellStyle(),i,j);

                cells.add(pdfpCell);
                j += colspan - 1;
            }
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
        Font font = excelReader.getWorkbook().getFontAt(cell.getCellStyle().getFontIndexAsInt());
        Paragraph paragraph = new Paragraph(getDataValueByKey(getKey(cell.getStringCellValue()))).setFont(f);
        if(font.getBold()){
            paragraph.setBold();
        }
        paragraph.setFontSize(size);
        return paragraph;
    }

    private Paragraph getValue(org.apache.poi.ss.usermodel.Cell cell, String language) {
        return getValue(cell,language,excelReader.getWorkbook().getFontAt(cell.getCellStyle().getFontIndexAsInt()).getFontHeightInPoints());
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
        if(!hasPic(cell)){
            return null;
        }
        Map<String, PictureData> picMap = ExcelPicUtil.getPicMap(excelReader.getWorkbook(),excelReader.getWorkbook().getSheetIndex(cell.getSheet().getSheetName()));
        PictureData picData = picMap.get(String.format("%s_%s",cell.getRowIndex(),cell.getColumnIndex()));
        if(picData == null){
            return null;
        }
        return picData.getData();
    }

    /**
     * 暂时只支持03 excel
     * @param cell
     * @return
     */
    private boolean hasPic(org.apache.poi.ss.usermodel.Cell cell) {
        if(excelReader.isXlsx()){
            return false;
        }else{
            if(cell.getSheet().getDrawingPatriarch() == null){
                return false;
            }
            return true;
        }
    }


    /**
     * 获取Excel的列宽像素
     * @param cell
     * @return
     */
    protected float getColumnWidth(org.apache.poi.ss.usermodel.Cell cell) {
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

    public int getDpi() {
        return dpi;
    }

    public ExcelObject setDpi(int dpi) {
        this.dpi = dpi;
        return this;
    }

    private String getDataValueByKey(CellValue cellValue){
        if(cellValue == null){
            return "";
        }

        if(StrUtil.isBlank(cellValue.getValue())){
            return "";
        }
        if(data == null){
            return cellValue.getValue();
        }
        if(!cellValue.isFlag()){
            return cellValue.getValue();
        }

        if(data instanceof Map){
            return Optional.ofNullable(((Map) data).get(cellValue.getValue())).orElse("").toString();
        }else{
            Object obj = ReflectUtil.invoke(data,StrUtil.upperFirstAndAddPre(cellValue.getValue(),"get"));
            return Optional.ofNullable(obj).orElse("").toString();
        }
    }

    private CellValue getKey(String text){
        boolean flag = false;
        CellValue cellValue = new CellValue();

        if(StrUtil.isBlank(text)){
            text = "";
        }

        if(text.length() > DELIMITER_PLACEHOLDER_START.length()){
            if(StrUtil.equals(DELIMITER_PLACEHOLDER_START,text.substring(0,DELIMITER_PLACEHOLDER_START.length()))){
                text = text.substring(DELIMITER_PLACEHOLDER_START.length(),text.length());
                flag = true;
            }
        }

        if(flag && text.length() >DELIMITER_PLACEHOLDER_END.length()){
            if(StrUtil.equals(DELIMITER_PLACEHOLDER_END,text.substring(text.length() -DELIMITER_PLACEHOLDER_END.length(),text.length()))){
                text = text.substring(0,text.length() - DELIMITER_PLACEHOLDER_END.length());
                flag = true;
            }
        }

        cellValue.setFlag(flag)
                .setText(text);
        return cellValue;
//        if(!flag){
//            return "";
//        }
//
//        return text;
    }

    public ExcelObject setData(Object data) {
        this.data = data;
        return this;
    }
}
