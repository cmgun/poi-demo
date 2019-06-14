package com.cmgun.excel;

import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.util.CollectionUtils;
import com.alibaba.excel.util.POITempFile;
import com.alibaba.excel.util.TypeUtil;
import com.alibaba.excel.util.WorkBookUtil;
import com.alibaba.excel.write.ExcelBuilder;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

/**
 * 根据Excel模板导出Excel，替换模板占位符。目前只为了兼容jxls-poi-jdk1.6的模板导出样式而已。
 * 目前只支持单sheet的模板导出操作。
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class ExcelJxlsTemplateBuilderImpl implements ExcelBuilder {

    /**
     * 专门用于处理Jxls模板的上下文
     */
    private JxlsWriteContext context;

    /**
     * 列模板表达式
     */
    private List<Expression> cellJexlExpressions = new ArrayList<>();

    /**
     * 列模板样式
     */
    private List<CellStyle> cellStyles = new ArrayList<>();

    /**
     * 列模板
     */
    private List<String> cellTemplates = new ArrayList<>();

    /**
     * 列模板对应的反射字段
     */
    private List<String> cellFieldNames = new ArrayList<>();

    /**
     * 模板最后一行的位置，默认为0
     */
    private int templateLastRowNum = 0;

    /**
     * jexlContext，用于模板
     */
    private JexlContext jexlContext = new MapContext();

    public ExcelJxlsTemplateBuilderImpl(InputStream templateInputStream, OutputStream out) {
        // 只读取模板的前N-1列作为模板头固定列，第N列为模板占位符替换位置
        try {
            //初始化时候创建临时缓存目录，用于规避POI在并发写bug
            POITempFile.createPOIFilesDirectory();
            /*
            inputSteam只能读取一次，但SXSSFWorkbook使用滑动窗口方式进行遍历，已经遍历过的Cell无法再写。
            因此这里先将 inputStream 封装为 XSSFWorkbook，读取最后一行后再清除最后一行，再生成 context 所需的 SXSSFWorkbook
             */
            // 解析模板excel的最后一行，读取占位符
            Workbook workbook = readLastRow(templateInputStream);

            // 创建写上下文
            context = new JxlsWriteContext(workbook, out);
            // 新建一个sheet
            Sheet sheet = new Sheet(1, templateLastRowNum);
            sheet.setSheetName("sheet1");
            sheet.setStartRow(templateLastRowNum + 1);
            sheet.setHeadLineMun(templateLastRowNum);
            context.currentSheet(sheet);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    public ExcelJxlsTemplateBuilderImpl(InputStream templateInputStream, OutputStream out, Map<String, Object> datas) {
        this(templateInputStream, out);
        // 初始化JexlContext
        initJexlContext(datas);
    }

    /**
     * 读取模板文件最后一行的数据，然后清除最后一行，将 workbook 返回，writeContext 可直接使用该对象
     *
     * @param tempInputStream 模板文件流
     * @throws Exception
     * @return 清除最后一行的workbook
     */
    private Workbook readLastRow(InputStream tempInputStream) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(tempInputStream);
        templateLastRowNum = workbook.getSheetAt(0).getLastRowNum();
        Row lastRow = workbook.getSheetAt(0).getRow(templateLastRowNum);
        // 解析最后一行，最后一行为占位符
        int colSize = lastRow.getLastCellNum();
        // Jexl解析引擎
        JexlEngine jexlEngine = new JexlEngine();
        for (int i = lastRow.getFirstCellNum(); i < colSize; i++) {
            Cell cell = lastRow.getCell(i);
            String cellTemplate = cell.getStringCellValue();
            // 添加列模板
            cellTemplates.add(cellTemplate);
            // 添加列模板表达式
            cellJexlExpressions.add(jexlEngine.createExpression(JxlsPlaceHolderUtils.convertPlaceHolder(cellTemplate)));
            // 添加列模板样式
            cellStyles.add(cell.getCellStyle());
            // 列模板解析，获取反射字段
            cellFieldNames.add(JxlsPlaceHolderUtils.getCellFieldName(cellTemplate));
        }
        // 清除最后一行
        workbook.getSheetAt(0).removeRow(lastRow);
        // 封装为 SXSSFWorkbook 后返回
        return new SXSSFWorkbook(workbook);
    }

    /**
     * 初始化JexlContext
     * @param datas 包含需要写入的数据和工具类
     */
    private void initJexlContext(Map<String, Object> datas) {
        // 获取解析工具类，除了 datas 以外的key
        Iterator<Map.Entry<String, Object>> iterator = datas.entrySet().iterator();
        while (iterator.hasNext()) {
            Map.Entry<String, Object> entry = iterator.next();
            if (!"datas".equals(entry.getKey())) {
                // 非数据内容，放入context后删除
                jexlContext.set(entry.getKey(), entry.getValue());
                iterator.remove();
            }
        }
    }

    /**
     * 只有这个方法在使用，从模板最后一行开始写数据
     *
     * @param data 需要写入的数据
     * @param startRow 暂时无用的字段
     */
    @Override
    public void addContent(List data, int startRow) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        // 从最后一行开始写
        templateLastRowNum = templateLastRowNum - 1;
        for (int i = 0; i < data.size(); i++) {
            int n = i + templateLastRowNum + 1;
            addOneRowOfDataToExcel(data.get(i), n);
        }
    }

    /**
     * 使用Jexl解析表达式
     * @param data
     */
    public void addContext(List data) {
        if (CollectionUtils.isEmpty(data)) {
            return;
        }
        // 从最后一行开始写
        templateLastRowNum = templateLastRowNum - 1;
        for (int i = 0; i < data.size(); i++) {
            int n = i + templateLastRowNum + 1;
            addOneRowOfDataToExcelWithJexl(data.get(i), n);
        }
    }

    /**
     * 添加一行数据到excel中，使用jexl进行模板字符串解析
     *
     * @param oneRowData 待添加的对象
     * @param n 待操作的行
     */
    private void addOneRowOfDataToExcelWithJexl(Object oneRowData, int n) {
        Row row = WorkBookUtil.createRow(context.getCurrentSheet(), n);
        // 添加当前beanMap到Context中
        jexlContext.set("c", oneRowData);
        for (int i = 0; i < cellJexlExpressions.size(); i++) {
            Object cellValue = JxlsPlaceHolderUtils.getCellValue(cellJexlExpressions.get(i), jexlContext);
            WorkBookUtil.createCell(row, i, cellStyles.get(i), cellValue, TypeUtil.isNum(cellValue));
        }
    }


    @Override
    public void addContent(List data, Sheet sheetParam) {
        throw new UnsupportedOperationException("暂不支持该用法");
    }

    @Override
    public void addContent(List data, Sheet sheetParam, Table table) {
        throw new UnsupportedOperationException("暂不支持该用法");
    }

    @Override
    public void merge(int firstRow, int lastRow, int firstCol, int lastCol) {
        throw new UnsupportedOperationException("暂不支持该用法");
    }

    @Override
    public void finish() {
        try {
            context.getWorkbook().write(context.getOutputStream());
            context.getWorkbook().close();
        } catch (IOException e) {
            throw new ExcelGenerateException("IO error", e);
        }
    }

    /**
     * 添加java对象到Excel中
     * @param oneRowData 待添加的对象
     * @param row 待操作的行
     */
    private void addJavaObjectToExcel(Object oneRowData, Row row) {
        BeanMap beanMap = BeanMap.create(oneRowData);
        for (int i = 0; i < cellTemplates.size(); i++) {
            String cellValue = getCellValue(cellTemplates.get(i), cellFieldNames.get(i), beanMap);
            WorkBookUtil.createCell(row, i, context.getCurrentContentStyle(), cellValue);
        }
    }

    /**
     * 获取单元格内容
     *
     * @param cellTemplate 单元格模板
     * @param fieldName 映射属性字段
     * @param beanMap bean
     * @return 需要写入的内容
     */
    private String getCellValue(String cellTemplate, String fieldName, BeanMap beanMap) {
        Object value = beanMap.get(fieldName);
        if (value == null) {
            // 无法找到对应字段，返回空
            return "";
        }
        String beanValue = value.toString();
        // 日期转换
        beanValue = JxlsPlaceHolderUtils.convertDateFormat(cellTemplate, beanValue);
        // 枚举常量转换
        beanValue = JxlsPlaceHolderUtils.convertConstantEnums(cellTemplate, beanValue);
        return beanValue;
    }

    /**
     * 添加一行新数据
     *
     * @param oneRowData 新数据
     * @param n 对应的行号
     */
    private void addOneRowOfDataToExcel(Object oneRowData, int n) {
        Row row = WorkBookUtil.createRow(context.getCurrentSheet(), n);
        addJavaObjectToExcel(oneRowData, row);
    }
}
