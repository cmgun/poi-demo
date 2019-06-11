package com.cmgun.poi;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

import java.util.Date;

public class Entity extends BaseRowModel {

    @ExcelProperty(value = {"id"})
    private long id;

    @ExcelProperty(value = {"msg"})
    private String msg;

    @ExcelProperty(value = {"msg1"})
    private String msg1;

    @ExcelProperty(value = {"msg2"})
    private String msg2;

    @ExcelProperty(value = {"createDate"}, format = "yyyy-MM-dd")
    private Date createDate;

    public Entity(long id, String msg) {
        this.id = id;
        this.msg = msg;
        this.msg1 = "msgmsgmsgmsgmsgaaaaaaaaaaa" + "1";
        this.msg2 = "msgmsgmsgmsgmsgaaaaabbbbbb"  + "2";
        this.createDate = new Date();
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
}
