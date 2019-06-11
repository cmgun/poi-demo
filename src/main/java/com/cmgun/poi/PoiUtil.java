package com.cmgun.poi;


import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;

import java.io.*;
import java.util.List;

public class PoiUtil {

    public static void export(String targetFileName, List<? extends BaseRowModel> datas) {
        OutputStream out = null;
        ExcelWriter writer = null;
        try {
            out = new FileOutputStream(targetFileName);
            writer = EasyExcelFactory.getWriter(out, ExcelTypeEnum.XLSX,true);
            Sheet sheet1 = new Sheet(1, 0, Entity.class, "第一个sheet", null);
            sheet1.setStartRow(0);
            writer.write(datas, sheet1);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 关闭资源
            if (writer != null) {
                writer.finish();
            }
            closeOutputStream(out);
        }
    }

    public static void export(String templateFileName, String targetFileName, List<? extends BaseRowModel> datas, int headLineNum) {
        OutputStream out = null;
        ExcelWriter writer = null;
        try {
            InputStream inputStream = getResourcesFileInputStream(templateFileName);
            out = new FileOutputStream(targetFileName);
            writer = EasyExcelFactory.getWriterWithTemp(inputStream,out, ExcelTypeEnum.XLSX,true);
            Sheet sheet1 = new Sheet(1, headLineNum, Entity.class, "第一个sheet", null);
            sheet1.setStartRow(headLineNum);
            writer.write(datas, sheet1);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 关闭资源
            if (writer != null) {
                writer.finish();
            }
            closeOutputStream(out);
        }
    }

    private static InputStream getResourcesFileInputStream(String fileName) {
        return Thread.currentThread().getContextClassLoader().getResourceAsStream("" + fileName);
    }

    private static void closeOutputStream(OutputStream out) {
        try {
            if (out != null) {
                out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
