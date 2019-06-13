package com.cmgun.excel;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.ExcelBuilderImpl;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

/**
 * 根据模板进行导出写入的Writer
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class ExcelTemplateWriter {

    private ExcelJxlsTemplateBuilderImpl excelBuilder;

    /**
     * 构造器
     *
     * @param templateInputStream 模板excel，包含占位符
     * @param outputStream 导出excel输出流
     * @param typeEnum 输出excel类型
     */
    public ExcelTemplateWriter(InputStream templateInputStream, OutputStream outputStream, ExcelTypeEnum typeEnum) {
        excelBuilder = new ExcelJxlsTemplateBuilderImpl(templateInputStream, outputStream, typeEnum);
    }

    /**
     * 写excel，不支持断续写入
     * @param data 待写入的数据
     */
    public void write(List<?> data) {
        excelBuilder.addContent(data, 0);
    }

    /**
     * 关闭IO流
     */
    public void finish() {
        excelBuilder.finish();
    }
}
