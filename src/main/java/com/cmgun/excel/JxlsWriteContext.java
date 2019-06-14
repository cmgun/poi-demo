package com.cmgun.excel;

import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.ExcelHeadProperty;
import com.alibaba.excel.metadata.Table;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.util.CollectionUtils;
import com.alibaba.excel.util.StyleUtil;
import com.alibaba.excel.util.WorkBookUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.OutputStream;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static com.alibaba.excel.util.StyleUtil.buildSheetStyle;

/**
 * Jxls模板写入上下文，根据 WriteContext 设计方式，增强对 workbook 的操作。
 *
 * @author chenqilin
 * @Date 2019/6/14
 */
public class JxlsWriteContext {

    /***
     * The sheet currently written
     */
    private Sheet currentSheet;

    /**
     * current param
     */
    private com.alibaba.excel.metadata.Sheet currentSheetParam;

    /**
     * The sheet currently written's name
     */
    private String currentSheetName;

    /**
     *
     */
    private Table currentTable;

    /**
     * Excel type
     */
    private ExcelTypeEnum excelType;

    /**
     * POI Workbook
     */
    private Workbook workbook;

    /**
     * Final output stream
     */
    private OutputStream outputStream;

    /**
     * Written form collection
     */
    private Map<Integer, Table> tableMap = new ConcurrentHashMap<Integer, Table>();

    /**
     * Cell default style
     */
    private CellStyle defaultCellStyle;

    /**
     * Current table head  style
     */
    private CellStyle currentHeadCellStyle;

    /**
     * Current table content  style
     */
    private CellStyle currentContentCellStyle;

    /**
     * the header attribute of excel
     */
    private ExcelHeadProperty excelHeadProperty;

    public JxlsWriteContext(Workbook workbook, OutputStream out) {
        this.outputStream = out;
        this.workbook = workbook;
        this.defaultCellStyle = StyleUtil.buildDefaultCellStyle(this.workbook);
    }

    /**
     * 设置当前Sheet
     * @param sheet
     */
    public void currentSheet(com.alibaba.excel.metadata.Sheet sheet) {
        if (null == currentSheetParam || currentSheetParam.getSheetNo() != sheet.getSheetNo()) {
            cleanCurrentSheet();
            currentSheetParam = sheet;
            try {
                this.currentSheet = workbook.getSheetAt(sheet.getSheetNo() - 1);
            } catch (Exception e) {
                this.currentSheet = WorkBookUtil.createSheet(workbook, sheet);
            }
            buildSheetStyle(currentSheet, sheet.getColumnWidthMap());
            this.initCurrentSheet(sheet);
        }
    }

    /**
     * 初始化当前Sheet
     * @param sheet
     */
    private void initCurrentSheet(com.alibaba.excel.metadata.Sheet sheet) {
        // Excel头部属性
        initExcelHeadProperty(sheet.getHead(), sheet.getClazz());
        // 表格样式
        initTableStyle(sheet.getTableStyle());
        // 表格头
        initTableHead();
    }

    /**
     * 清空当前Sheet内容
     */
    private void cleanCurrentSheet() {
        this.currentSheet = null;
        this.currentSheetParam = null;
        this.excelHeadProperty = null;
        this.currentHeadCellStyle = null;
        this.currentContentCellStyle = null;
        this.currentTable = null;
    }

    /**
     * init excel header
     *
     * @param head
     * @param clazz
     */
    private void initExcelHeadProperty(List<List<String>> head, Class<? extends BaseRowModel> clazz) {
        if (head != null || clazz != null) {
            this.excelHeadProperty = new ExcelHeadProperty(clazz, head);
        }
    }

    /**
     * 初始化表格头
     */
    public void initTableHead() {
        if (null != excelHeadProperty && !CollectionUtils.isEmpty(excelHeadProperty.getHead())) {
            int startRow = currentSheet.getLastRowNum();
            if (startRow > 0) {
                startRow += 4;
            }else {
                startRow = currentSheetParam.getStartRow();
            }
            addMergedRegionToCurrentSheet(startRow);
            int i = startRow;
            for (; i < this.excelHeadProperty.getRowNum() + startRow; i++) {
                Row row = WorkBookUtil.createRow(currentSheet, i);
                addOneRowOfHeadDataToExcel(row, this.excelHeadProperty.getHeadByRowNum(i - startRow));
            }
        }
    }

    /**
     * 合并单元格
     * @param startRow
     */
    private void addMergedRegionToCurrentSheet(int startRow) {
        for (com.alibaba.excel.metadata.CellRange cellRangeModel : excelHeadProperty.getCellRangeModels()) {
            currentSheet.addMergedRegion(new CellRangeAddress(cellRangeModel.getFirstRow() + startRow,
                    cellRangeModel.getLastRow() + startRow,
                    cellRangeModel.getFirstCol(), cellRangeModel.getLastCol()));
        }
    }

    /**
     * 添加一行表格头数据
     * @param row
     * @param headByRowNum
     */
    private void addOneRowOfHeadDataToExcel(Row row, List<String> headByRowNum) {
        if (headByRowNum != null && headByRowNum.size() > 0) {
            for (int i = 0; i < headByRowNum.size(); i++) {
                WorkBookUtil.createCell(row, i, getCurrentHeadCellStyle(), headByRowNum.get(i));
            }
        }
    }

    /**
     * 初始化表格样式
     * @param tableStyle
     */
    private void initTableStyle(com.alibaba.excel.metadata.TableStyle tableStyle) {
        if (tableStyle != null) {
            this.currentHeadCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableHeadFont(),
                    tableStyle.getTableHeadBackGroundColor());
            this.currentContentCellStyle = StyleUtil.buildCellStyle(this.workbook, tableStyle.getTableContentFont(),
                    tableStyle.getTableContentBackGroundColor());
        }
    }
    

    public ExcelHeadProperty getExcelHeadProperty() {
        return this.excelHeadProperty;
    }

    public Sheet getCurrentSheet() {
        return currentSheet;
    }

    public void setCurrentSheet(Sheet currentSheet) {
        this.currentSheet = currentSheet;
    }

    public String getCurrentSheetName() {
        return currentSheetName;
    }

    public void setCurrentSheetName(String currentSheetName) {
        this.currentSheetName = currentSheetName;
    }

    public ExcelTypeEnum getExcelType() {
        return excelType;
    }

    public void setExcelType(ExcelTypeEnum excelType) {
        this.excelType = excelType;
    }

    public OutputStream getOutputStream() {
        return outputStream;
    }

    public CellStyle getCurrentHeadCellStyle() {
        return this.currentHeadCellStyle == null ? defaultCellStyle : this.currentHeadCellStyle;
    }

    public CellStyle getCurrentContentStyle() {
        return this.currentContentCellStyle;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public com.alibaba.excel.metadata.Sheet getCurrentSheetParam() {
        return currentSheetParam;
    }

    public void setCurrentSheetParam(com.alibaba.excel.metadata.Sheet currentSheetParam) {
        this.currentSheetParam = currentSheetParam;
    }

    public Table getCurrentTable() {
        return currentTable;
    }

    public void setCurrentTable(Table currentTable) {
        this.currentTable = currentTable;
    }
}
