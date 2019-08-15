package com.cmgun.poi;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;
import com.cmgun.excel.extend.CellStyleEnum;
import com.cmgun.excel.extend.ExcelCellStyle;

import java.math.BigDecimal;
import java.util.Date;

public class Entity extends BaseRowModel {

    @ExcelProperty(index = 0, value = {"id"})
    private long id;

    @ExcelProperty(index = 1, value = {"msg"})
    private String msg;

    @ExcelProperty(index = 2, value = {"msg1"})
    private String msg1;

    @ExcelProperty(index = 3, value = {"msg2"})
    private String msg2;

    @ExcelProperty(index = 4, value = {"money"})
    @ExcelCellStyle(cellStyle = CellStyleEnum.MONEY, format = "#,##0.00")
    private BigDecimal money;

    @ExcelProperty(index = 5, value = {"createDate"}, format = "yyyy-MM-dd")
    @ExcelCellStyle(cellStyle = CellStyleEnum.DATE, format = "yyyy-MM-dd")
    private Date createDate;

    private String strCreateDate;

    private String constantVal;

    public Entity() {
    }

    public Entity(long id, String msg, String constantVal) {
        this.id = id;
        this.msg = msg;
        this.msg1 = "msgmsgmsgmsgmsgaaaaaaaaaaa" + "1";
        this.msg2 = "msgmsgmsgmsgmsgaaaaabbbbbb"  + "2";
        this.createDate = new Date();
        this.money = new BigDecimal("1000000.12");
        this.strCreateDate = "2020-12-12 12:12:12";
        this.constantVal = constantVal;
    }

    public long getId() {
        return id;
    }

    public void setId(long id) {
        this.id = id;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public Date getCreateDate() {
        return createDate;
    }

    public void setCreateDate(Date createDate) {
        this.createDate = createDate;
    }

    public String getMsg1() {
        return msg1;
    }

    public void setMsg1(String msg1) {
        this.msg1 = msg1;
    }

    public String getMsg2() {
        return msg2;
    }

    public void setMsg2(String msg2) {
        this.msg2 = msg2;
    }

    public String getStrCreateDate() {
        return strCreateDate;
    }

    public void setStrCreateDate(String strCreateDate) {
        this.strCreateDate = strCreateDate;
    }

    public String getConstantVal() {
        return constantVal;
    }

    public void setConstantVal(String constantVal) {
        this.constantVal = constantVal;
    }

    public BigDecimal getMoney() {
        return money;
    }

    public void setMoney(BigDecimal money) {
        this.money = money;
    }

    @Override
    public String toString() {
        return "Entity{" +
                "id=" + id +
                ", msg='" + msg + '\'' +
                ", msg1='" + msg1 + '\'' +
                ", msg2='" + msg2 + '\'' +
                ", money=" + money +
                ", createDate=" + createDate +
                ", strCreateDate='" + strCreateDate + '\'' +
                ", constantVal='" + constantVal + '\'' +
                '}';
    }
}
