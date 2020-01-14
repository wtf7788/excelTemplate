package com.tf.print.template.util;

import com.tf.print.template.excel.ExcelExReader;

import java.io.File;

/**
 * @ClassName ExcelTemplateUtil
 * @Description TODO
 * @Author kyjonny
 * @Date 3/1/2020 10:04 上午
 **/
public class ExcelTemplateUtil {

    public static ExcelExReader getReader(File bookFile, int sheetIndex){
        return new ExcelExReader(bookFile, sheetIndex);
    }

    public static ExcelExReader getReader(String bookFilePath){
        return new ExcelExReader(bookFilePath,0);
    }

    public static ExcelExReader getReader(String bookFilePath, int sheetIndex){
        return new ExcelExReader(bookFilePath, sheetIndex);
    }
}
