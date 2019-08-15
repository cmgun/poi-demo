package com.cmgun.excel.extend;

import java.lang.annotation.*;

/**
 * 扩展单元格样式
 *
 * @author chenqilin
 * @date 2019/8/15
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelCellStyle {

    /**
     * 单元格样式
     * @return
     */
    CellStyleEnum cellStyle();

    /**
     * 格式化内容
     * @return
     */
    String format() default "";
}
