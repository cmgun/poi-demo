package com.cmgun.excel.footer;

import java.util.ArrayList;
import java.util.List;

/**
 * 记录模板footer一行的信息
 * @author chenqilin
 * @Date 2019/6/14
 */
public class FooterRow {

    /**
     * 行号，距离最后一条数据模板的数据的行距离
     */
    private int footerRowNum;

    private List<FooterCell> cells = new ArrayList<>();

    public FooterRow(int footerRowNum) {
        this.footerRowNum = footerRowNum;
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
