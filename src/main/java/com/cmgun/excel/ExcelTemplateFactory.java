package com.cmgun.excel;

import com.alibaba.excel.EasyExcelFactory;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * Excel操作工厂类
 * 增强功能：
 * 1. 支持从模板excel中提取占位符进行javaBean映射导出
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class ExcelTemplateFactory extends EasyExcelFactory {

    /**
     * 获取读取 Jxls-poi-jdk1.6 模板的excel writer，目前只支持 .xlsx 后缀模板文件
     *
     * @param temp 模板文件流
     * @param outputStream 导出目标文件流
     * @return excel writer
     */
    public static ExcelTemplateWriter getWriterWithTemp(InputStream temp, OutputStream outputStream) {
        return new ExcelTemplateWriter(temp, outputStream);
    }
}
