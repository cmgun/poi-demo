package com.cmgun.excel.template;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

/**
 * 根据模板进行导出写入的Writer
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class ExcelTemplateWriter {

    private ExcelJxlsTemplateBuilderImpl excelBuilder;

    public ExcelTemplateWriter(InputStream templateInputStream, OutputStream out, Map<String, Object> datas) {
        excelBuilder = new ExcelJxlsTemplateBuilderImpl(templateInputStream, out, datas);
    }

    /**
     * 写excel，不支持断续写入
     * @param data 待写入的数据
     * @param startRow
     */
    public void write(List<?> data, int startRow) {
        excelBuilder.addContent(data, 0);
    }

    public void write(Map<String, Object> data) {
        excelBuilder.addContext((List) data.get("datas"));
    }

    /**
     * 关闭IO流
     */
    public void finish() {
        excelBuilder.finish();
    }
}
