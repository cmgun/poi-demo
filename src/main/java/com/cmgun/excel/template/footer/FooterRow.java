package com.cmgun.excel.template.footer;

import org.apache.commons.jexl2.JexlContext;
import org.apache.commons.jexl2.JexlEngine;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.util.ArrayList;
import java.util.List;

/**
 * 记录模板footer一行的信息
 *
 * @author chenqilin
 * @Date 2019/6/14
 */
public class FooterRow {

    /**
     * 行号，距离最后一条数据模板的数据的行距离
     */
    private int footerRowNum;

    private List<FooterCell> cells = new ArrayList<>();

    public FooterRow(Row row, int footerRowNum, JexlEngine jexlEngine, JexlContext jexlContext) {
        this.footerRowNum = footerRowNum;
        // 解析cell内容
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            // 不存空格
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            cells.add(new FooterCell(cell, jexlEngine, jexlContext));
        }
    }

    public int getFooterRowNum() {
        return footerRowNum;
    }

    public void setFooterRowNum(int footerRowNum) {
        this.footerRowNum = footerRowNum;
    }

    public List<FooterCell> getCells() {
        return cells;
    }

    public void setCells(List<FooterCell> cells) {
        this.cells = cells;
    }
}
