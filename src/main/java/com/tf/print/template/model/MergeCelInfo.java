package com.tf.print.template.model;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * @ClassName MergeCelInfos
 * @Description 合并单元格信息
 * @Author kyjonny
 * @Date 14/1/2020 9:40 上午
 **/
@Data
@Accessors(chain = true)
public class MergeCelInfo {


    private int row =1;


    private int col =1;


    private boolean skipRow = true;

    private int step =1;

}
