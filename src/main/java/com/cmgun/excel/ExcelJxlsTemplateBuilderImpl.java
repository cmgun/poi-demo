package com.cmgun.excel;

import com.alibaba.excel.exception.ExcelGenerateException;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.util.CollectionUtils;
import com.alibaba.excel.util.POITempFile;
import com.alibaba.excel.util.TypeUtil;
import com.alibaba.excel.util.WorkBookUtil;
import com.alibaba.excel.write.ExcelBuilder;
import com.cmgun.excel.footer.FooterCell;
import com.cmgun.excel.footer.FooterRow;
import net.sf.cglib.beans.BeanMap;
import org.apache.commons.jexl2.Expression;
import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.commons.jexl2.MapContext;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
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
     * 模板的footers
     */
    private List<FooterRow> footers = new ArrayList<>();

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
     * 占位符行的位置，默认为0
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
            因此这里先将 inputStream 封装为 XSSFWorkbook，读取占位符行后再删除对应的行，最后生成 context 所需的 SXSSFWorkbook
             */
            // 解析模板excel，读取占位符
            Workbook workbook = readCellTemplate(templateInputStream);

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
    private Workbook readCellTemplate(InputStream tempInputStream) throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook(tempInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);
        // 是否已经读完所有模板头，即占位符开始前的位置
        boolean hasReadAllHeadRows = false;
        // footer部分模板行数
        int footerRows = 0;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (!hasReadAllHeadRows) {
                Cell cell = row.getCell(0);
                // 还未读完模板头部，判断第一个cell是否包含占位符
                if (cell != null && JxlsPlaceHolderUtils.isPlaceHolderCell(cell.getStringCellValue())) {
                    // 包含占位符，模板头读取结束
                    hasReadAllHeadRows = true;
                    // 数据模板解析
                    initCellTemplates(row);
                    // 清除该行
                    sheet.removeRow(row);
                }
            } else{
                // 已经读完头部数据，剩下的行属于footer部分，记录后从sheet中清除
                // 不能直接存Row对象，会被disconnect，因此只记录cell文字和格式
                if (row != null) {
                    footers.add(new FooterRow(row, footerRows));
                    sheet.removeRow(row);
                }
                footerRows++;
            }
        }
        // 封装为 SXSSFWorkbook 后返回
        return new SXSSFWorkbook(workbook);
    }

    /**
     * 初始化数据读取模板
     * @param row 模板行
     */
    private void initCellTemplates(Row row) {
        // 数据模板开始写入的行数
        templateLastRowNum = row.getRowNum();
        int colSize = row.getLastCellNum();
        // Jexl解析引擎
        JexlEngine jexlEngine = new JexlEngine();
        for (int i = row.getFirstCellNum(); i < colSize; i++) {
            Cell cell = row.getCell(i);
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
        int lastRowNum = 0;
        for (int i = 0; i < data.size(); i++) {
            lastRowNum = i + templateLastRowNum + 1;
            addOneRowOfDataToExcelWithJexl(data.get(i), lastRowNum);
        }
        // 写footer
        for (int i = 0; i < footers.size(); i++) {
            int n = i + lastRowNum + 1;
            addOneRowOfFooterToExcelWithJexl(footers.get(i), n, data.size());
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

    /**
     * 添加一行footer到excel
     * @param footerRow 待添加的row，如果有占位符需要替换
     * @param n 待写入的行数
     * @param dataSize 模板数据的长度
     */
    private void addOneRowOfFooterToExcelWithJexl(FooterRow footerRow, int n, int dataSize) {
        if (footerRow == null) {
            return;
        }
        int rowNum = n + footerRow.getFooterRowNum();
        Row row = WorkBookUtil.createRow(context.getCurrentSheet(), rowNum);
        for (int i = 0; i < footerRow.getCells().size(); i++) {
            FooterCell footerCell = footerRow.getCells().get(i);

            JxlsPlaceHolderUtils.convertFooterCell(footerCell, dataSize);
            // 设置footer cell
            Cell newCell = row.createCell(footerCell.getCellNum());
            newCell.setCellStyle(footerCell.getCellStyle());
            newCell.setCellFormula(footerCell.getCellFormula());
            newCell.setCellValue(footerCell.getCellValue());
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
