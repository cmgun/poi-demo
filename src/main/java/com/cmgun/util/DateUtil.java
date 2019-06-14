package com.cmgun.util;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * 模拟项目中的日期工具类
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class DateUtil {

    public static String convertToDate(String dateStr, String format) {
        if (dateStr == null || dateStr.length() == 0) {
            return null;
        }
        DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        Date date = null;
        try {
            date = dateFormat.parse(dateStr);
            DateFormat dateFormat1 = new SimpleDateFormat(format);
            return dateFormat1.format(date);
        } catch (ParseException e) {
            System.err.println(e);
        }
        return "";
    }
}
