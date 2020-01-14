package com.tf.print.template.model;

import com.tf.print.template.exception.TemplateException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import java.io.Serializable;

/**
 * @ClassName PositionModel
 * @Description 位置模型
 * @Author kyjonny
 * @Date 7/1/2020 10:01 上午
 **/
public class PositionModel implements Serializable {

    /**
     * 行索引
     */
    private int rowIndex;

    /**
     * 列索引
     */
    private int colIndex;

    /**
     * 当前单元格的类型
     */
    private Cell cell;

    /**
     * 列合并数量
     */
    private int colSpan;

    /**
     * 行合并数量
     */
    private int rowSpan;

    /**
     * 跳过当前单元格
     */
    private boolean skip = false;

    private PositionModel() {
    }

    public static PositionModel createPositionModel(Cell cell){
        PositionModel positionModel = new PositionModel();
        positionModel.colIndex = cell.getColumnIndex();
        positionModel.rowIndex = cell.getRowIndex();
        positionModel.cell = cell;
        return positionModel;
    }

    public int getColSpan() {
        return colSpan;
    }

    public void setColSpan(int colSpan) {
        this.colSpan = colSpan;
    }

    public int getRowSpan() {
        return rowSpan;
    }

    public void setRowSpan(int rowSpan) {
        this.rowSpan = rowSpan;
    }

    public boolean isSkip() {
        return skip;
    }

    public void setSkip(boolean skip) {
        this.skip = skip;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public int getColIndex() {
        return colIndex;
    }

    public Cell getCell() {
        return cell;
    }

    public String getValue(){
        String result = "";
        switch (cell.getCellType()){
            case BLANK:
                break;
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)){
                    result = cell.getDateCellValue().toString();
                }else{
                    result = String.valueOf(cell.getNumericCellValue());
                }
                break;
            case STRING:
                result = cell.getStringCellValue();
                break;
            case _NONE:
                break;
            default:
                throw new TemplateException("找不到对应的模版字段类型");
        }
        return result;
    }


    public int getBorderTop(){
        return cell.getCellStyle().getBorderTop().getCode();
    }

    public int getBorderBottom(){
        if(rowSpan >1){
            return cell.getCellStyle().getBorderTop().getCode();
        }
        return cell.getCellStyle().getBorderBottom().getCode();
    }

    public int getBorderLeft(){
        return cell.getCellStyle().getBorderLeft().getCode();
    }

    public int getBorderRight(){
        if(colSpan >1){
            return cell.getCellStyle().getBorderLeft().getCode();
        }
        return cell.getCellStyle().getBorderRight().getCode();
    }

}
