package com.cmgun.util;

/**
 * 模拟项目中的TranslateUtil
 *
 * @author chenqilin
 * @Date 2019/6/13
 */
public class TranslateUtil {

    public static String getConstantName(String constantType, String constantValue) {
        if ("0".equals(constantValue)) {
            return "这是个值为0的枚举";
        } else if ("1".equals(constantValue)) {
            return "这是个值为1的枚举";
        }
        return "没有定义的枚举";
    }
}
