package com.cmgun.excel.footer;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;

/**
 * Footer Row 的单元格
 * @author chenqilin
 * @Date 2019/6/14
 */
public class FooterCell {

    /**
     * 列数
     */
    private int cellNum;

    private CellStyle cellStyle;

    private CellFormat cellFormat;

    private CellValue cellValue;

    public FooterCell(int cellNum, CellStyle cellStyle, CellFormat cellFormat, CellValue cellValue) {
        this.cellNum = cellNum;
        this.cellStyle = cellStyle;
        this.cellFormat = cellFormat;
        this.cellValue = cellValue;
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

    public CellFormat getCellFormat() {
        return cellFormat;
    }

    public void setCellFormat(CellFormat cellFormat) {
        this.cellFormat = cellFormat;
    }

    public CellValue getCellValue() {
        return cellValue;
    }

    public void setCellValue(CellValue cellValue) {
        this.cellValue = cellValue;
    }
}
