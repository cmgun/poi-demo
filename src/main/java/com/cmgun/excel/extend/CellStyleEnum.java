package com.cmgun.excel.extend;

/**
 * 自定义单元格样式枚举类
 *
 * @author chenqilin
 * @date 2019/8/15
 */
public enum CellStyleEnum {

    /**
     * 日期
     */
    DATE,
    /**
     * 金额
     */
    MONEY;

    /**
     * 枚举中是否包含指定对象
     *
     * @param o
     * @return
     */
    public static boolean contains(Object o) {
        for (CellStyleEnum cellStyleEnum : CellStyleEnum.values()) {
            if (cellStyleEnum.equals(o)) {
                return true;
            }
        }
        return false;
    }
}
