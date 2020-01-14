package com.tf.print.template.excel;

import cn.hutool.core.io.FileUtil;
import cn.hutool.poi.excel.ExcelBase;
import cn.hutool.poi.excel.WorkbookUtil;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import java.io.File;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * excel扩展读取工具
 */
public class ExcelExReader extends ExcelBase<ExcelExReader> {

    private Map<String, String> headerAlias;

    boolean read07;

    private float countWidthInPixels = 0F;

    public ExcelExReader(String excelFilePath, int sheetIndex) {
        this(FileUtil.file(excelFilePath), sheetIndex);
    }

    public ExcelExReader(File bookFile, int sheetIndex) {
        this(WorkbookUtil.createBook(bookFile), sheetIndex);
    }

    public ExcelExReader(InputStream bookStream, int sheetIndex, boolean closeAfterRead) {
        this(WorkbookUtil.createBook(bookStream, closeAfterRead), sheetIndex);
    }

    public ExcelExReader(InputStream bookStream, String sheetName, boolean closeAfterRead) {
        this(WorkbookUtil.createBook(bookStream, closeAfterRead), sheetName);
    }

    public ExcelExReader(Workbook book, String sheetName) {
        this(book.getSheet(sheetName));
    }

    public ExcelExReader(Workbook book, int sheetIndex) {
        this(book.getSheetAt(sheetIndex));
    }

    public ExcelExReader(Sheet sheet) {
        super(sheet);
        this.headerAlias = new HashMap();
        read07 = this.sheet instanceof XSSFSheet;
    }

    /**
     * 获取excel表格总宽度
     * @return
     */
    public synchronized float getCountWidthInPixels() {
        if(countWidthInPixels == 0){
            for(int i=0 ;i<getColumnCount();i++){
                countWidthInPixels += sheet.getColumnWidthInPixels(i);
            }
        }
        return countWidthInPixels;
    }
}
