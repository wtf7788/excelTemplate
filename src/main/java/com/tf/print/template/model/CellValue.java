package com.tf.print.template.model;

import lombok.Data;
import lombok.experimental.Accessors;
import java.io.Serializable;

/**
 * @ClassName CellValue
 * @Description TODO
 * @Author kyjonny
 * @Date 22/1/2020 10:00 下午
 **/
@Data
@Accessors(chain = true)
public class CellValue implements Serializable {

    private String text;

    /**
     * 是否是填充值类型
     */
    private boolean flag;

    public CellValue() {
    }

    public CellValue(String text, boolean flag) {
        this.text = text;
        this.flag = flag;
    }

    public String getValue(){
        return text;
    }
}
