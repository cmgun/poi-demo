package com.cmgun.excel;

import com.cmgun.util.DateUtil;
import com.cmgun.util.TranslateUtil;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 兼容jxls-poi-jdk1.6的模板解析方式
 * 目前只支持一下占位符的解析：
 * 1. 属性字段，即在${...}中以c.开头的属性字段，
 *    如 ${c.field}，${translateUtil.getConstantName('Field', c.field)}，${dateUtil.convertToDate(c.field, 'yyyyMMdd')}
 * 2. translateUtil.getConstantName，常量转换方法
 * 3. dateUtil.convertToDate，String类型日期的格式转换
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class JxlsPlaceHolderUtils {

    /**
     * cell value的field名称，以c.*开头，)} 或 } 或 , 结尾，*为field名称
     */
    private static final String CELL_VALUE_PATTERN = ".*c\\.([^),]*),?.*\\)?}";

    /**
     * cell里使用了工具类的参数，只处理TranslateUtil和DateUtil两个工具类
     */
    private static final String UTILS_VALUE_PATTERN = ".*'([^']*)'.*";

    /**
     * 日期转换工具类名称，处理String类型的日期转换为特定格式
     */
    private static final String DATE_UTIL = "dateUtil";

    /**
     * 业务常量转换工具类，处理模板中的枚举值转换
     */
    private static final String TRANSLATE_UTIL = "translateUtil.getConstantName";

    /**
     * 查找模板的反射field
     * @param cellValue 单元格内容
     * @return 匹配的field
     */
    public static String getCellFieldName(String cellValue) {
        Pattern pattern = Pattern.compile(CELL_VALUE_PATTERN);
        Matcher matcher = pattern.matcher(cellValue);
        if (matcher.matches()) {
            // 只会匹配一次
            return matcher.group(1);
        }
        return null;
    }

    /**
     * 转换字符串时间格式
     * @param cellTemplate 模板
     * @param beanValue 对应的bean的field值
     * @return 转换后的格式
     */
    public static String convertDateFormat(String cellTemplate, String beanValue) {
        if (!cellTemplate.contains(DATE_UTIL)) {
            // 不包含日期转换工具类，不用处理
            return beanValue;
        }
        // 进行内容转换
        // 获取单元格模板中的日期格式参数
        Pattern pattern = Pattern.compile(UTILS_VALUE_PATTERN);
        Matcher matcher = pattern.matcher(cellTemplate);
        if (matcher.matches()) {
            // 获取引号内的日期格式，模板中的日期格式为输出的指定日期格式
            return DateUtil.convertToFormatDateStr(beanValue, matcher.group(1));
        }
        // 模板内没有日期格式，不进行格式化
        return beanValue;
    }

    /**
     * 转换业务枚举值（实际是常量）
     * @param cellTemplate 模板
     * @param beanValue 对应的bean的field值
     * @return 转换后的格式
     */
    public static String convertConstantEnums(String cellTemplate, String beanValue) {
        if (!cellTemplate.contains(TRANSLATE_UTIL)) {
            // 不包含枚举转换工具类，不用处理
            return beanValue;
        }
        // 进行内容转换
        // 获取单元格模板中的参数
        Pattern pattern = Pattern.compile(UTILS_VALUE_PATTERN);
        Matcher matcher = pattern.matcher(cellTemplate);
        if (matcher.matches()) {
            // 获取引号内的参数
            return TranslateUtil.getConstantName(matcher.group(1), beanValue);
        }
        // 模板内没有枚举参数，不处理
        return beanValue;
    }
}
