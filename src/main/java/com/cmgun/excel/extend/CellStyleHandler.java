package com.cmgun.excel.extend;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import com.cmgun.util.DateUtil;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * 单元格样式扩展
 * 初始化时读取CellStyle，写入时根据初始化CellStyle进行单元格写入
 *
 * @author chenqilin
 * @date 2019/8/15
 */
public class CellStyleHandler<T extends Class<? extends BaseRowModel>> implements WriteHandler {

    /**
     * 导出的JavaBean对应的类
     */
    private T rowModel;

    /**
     * 单元格样式配置
     */
    private Map<Integer, ExcelCellStyle> styleConfigs = new HashMap<>();

    /**
     * 单元格样式
     */
    private Map<Integer, CellStyle> cellStyles = new HashMap<>();

    /**
     * 表头行数
     */
    private int headRows = 0;

    public CellStyleHandler(T rowModel) {
        this.rowModel = rowModel;
        init(rowModel);
    }

    /**
     * 初始化自定义格式
     */
    private void init(T rowModel) {
        Field[] declaredFields = rowModel.getDeclaredFields();
        int colNum = 0;
        for (Field field : declaredFields) {
            ExcelCellStyle excelCellStyle = field.getAnnotation(ExcelCellStyle.class);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelCellStyle != null && excelProperty != null) {
                styleConfigs.put(colNum, excelCellStyle);
                // 记录最大表头行数
                headRows = Math.max(excelProperty.value().length, headRows);
            }
            colNum++;
        }
    }

    @Override
    public void sheet(int i, Sheet sheet) {
        if (i != 1) {
            return;
        }
        // 只初始化一次
        Workbook workbook = sheet.getWorkbook();
        for (Map.Entry<Integer, ExcelCellStyle> entry : styleConfigs.entrySet()) {
            // 生成相应的CellStyle
            CellStyle cellStyle = null;
            if (CellStyleEnum.contains(entry.getValue().cellStyle())) {
                cellStyle = workbook.createCellStyle();
                DataFormat format = workbook.createDataFormat();
                cellStyle.setDataFormat(format.getFormat(entry.getValue().format()));
            }

            // 不为空则添加
            if (cellStyle != null) {
                cellStyles.put(entry.getKey(), cellStyle);
            }
        }
    }

    @Override
    public void row(int i, Row row) {
        // do nothing
    }

    @Override
    public void cell(int i, Cell cell) {
        if (!cellStyles.containsKey(i) || cell.getRowIndex() <= headRows - 1) {
            return;
        }
        // 有自定义样式
        CellStyle cellStyle = cellStyles.get(i);
        ExcelCellStyle excelCellStyle = styleConfigs.get(i);
        // cell value改为对应类型
        if (CellStyleEnum.DATE.equals(excelCellStyle.cellStyle())) {
            // 日期类型
            cell.setCellValue(DateUtil.format(cell.getStringCellValue(), excelCellStyle.format()));
        } else if (CellStyleEnum.MONEY.equals(excelCellStyle.cellStyle())) {
            // 金额类型
            cell.setCellValue(cell.getNumericCellValue());
        }
        cell.setCellStyle(cellStyle);
    }


}
