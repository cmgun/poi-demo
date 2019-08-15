package com.cmgun.util;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * @author chenqilin
 * @Date 2019/8/15
 */
public class CellStyleUtil {

//    public static CellStyle getDateFormatStyle(Workbook workbook, String format) {
//        HSSFCellStyle cellStyle = workbook.createCellStyle();
//        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(format));
//        return cellStyle;
//    }


    public static CellStyle getCurrencyStyle(String formatString) {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        HSSFDataFormat format = workbook.createDataFormat();
        // "¥#,##0"
        cellStyle.setDataFormat(format.getFormat(formatString));
        return cellStyle;
    }
}
