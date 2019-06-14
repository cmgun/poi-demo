package com.cmgun.poi;


import com.alibaba.excel.EasyExcelFactory;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.metadata.BaseRowModel;
import com.alibaba.excel.metadata.Sheet;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.cmgun.excel.ExcelTemplateFactory;
import com.cmgun.excel.ExcelTemplateWriter;
import org.apache.poi.util.IOUtils;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;

public class PoiUtil {

    /**
     * EasyExcel 无模板导出
     *
     * @param targetFileName
     * @param datas
     */
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

    /**
     * EasyExcel 模板Excel导出，只是拿模板样式不替换占位符
     *
     * @param templateFileName
     * @param targetFileName
     * @param datas
     * @param headLineNum
     */
    public static void export(String templateFileName, String targetFileName, List<? extends BaseRowModel> datas, int headLineNum) {
        OutputStream out = null;
        ExcelWriter writer = null;
        InputStream inputStream = null;
        try {
            inputStream = getResourcesFileInputStream(templateFileName);
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
            IOUtils.closeQuietly(out);
            IOUtils.closeQuietly(inputStream);
        }
    }

    /**
     * 读取excel模板，替换占位符内容，兼容Jxls模板读取方式
     */
    public static void exportForJxlsTemp(String templateFileName, String targetFileName, List datas) {
        OutputStream out = null;
        ExcelTemplateWriter writer = null;
        InputStream inputStream = null;
        try {
            inputStream = getResourcesFileInputStream(templateFileName);
            out = new FileOutputStream(targetFileName);
            writer = ExcelTemplateFactory.getWriterWithTemp(inputStream, out);
            writer.write(datas);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            // 关闭资源
            if (writer != null) {
                writer.finish();
            }
            IOUtils.closeQuietly(out);
            IOUtils.closeQuietly(inputStream);
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
