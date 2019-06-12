package excel;

import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.OutputStream;

/**
 * Excel操作工厂类
 * 增强功能：
 * 1. 支持从模板excel中提取占位符进行javaBean映射导出
 *
 * @author cmgun
 */
public class ExcelFactory extends EasyExcelFactory {

    public static ExcelTemplateWriter getTemplateWriter(OutputStream outputStream) {
        return new ExcelTemplateWriter(outputStream, ExcelTypeEnum.XLSX, true);
    }
}
