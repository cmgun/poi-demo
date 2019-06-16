package com.cmgun.excel.footer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;

/**
 * Footer Row 的单元格
 *
 * @author chenqilin
 * @Date 2019/6/14
 */
public class FooterCell {

    /**
     * 列数
     */
    private int cellNum;

    /**
     * 单元格样式
     */
    private CellStyle cellStyle;

    /**
     * 单元格式化表达式
     */
    private String cellFormula;

    /**
     * 单元格内容
     */
    private String cellValue;

    public FooterCell(Cell cell) {
        this.cellNum = cell.getColumnIndex();
        this.cellStyle = cell.getCellStyle();
        this.cellValue = cell.getStringCellValue();
    }

    public int getCellNum() {
        return cellNum;
    }

    public void setCellNum(int cellNum) {
        this.cellNum = cellNum;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getCellFormula() {
        return cellFormula;
    }

    public void setCellFormula(String cellFormula) {
        this.cellFormula = cellFormula;
    }

    public String getCellValue() {
        return cellValue;
    }

    public void setCellValue(String cellValue) {
        this.cellValue = cellValue;
    }
}
