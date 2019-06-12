package excel;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.ExcelBuilderImpl;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * 根据模板进行导出写入的Writer
 *
 * @author cmgun
 */
public class ExcelTemplateWriter {

    /**
     * 构造器
     * @param templateInputStream 模板excel，包含占位符
     * @param outputStream 导出excel输出流
     * @param typeEnum 输出excel类型
     */
    public ExcelTemplateWriter(InputStream templateInputStream, OutputStream outputStream, ExcelTypeEnum typeEnum) {
        excelBuilder = new ExcelBuilderImpl(templateInputStream,outputStream, typeEnum, false);
    }

}
