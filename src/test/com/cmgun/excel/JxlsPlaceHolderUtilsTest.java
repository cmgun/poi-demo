package com.cmgun.excel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.junit.Assert;
import org.junit.Test;

import java.util.regex.Matcher;

/**
 * 模板占位符转换测试
 *
 * @author chenqilin
 * @Date 2019/6/13
 */

public class JxlsPlaceHolderUtilsTest {

    @Test
    public void testConvertPlaceHolder() {
        String pattern1 = "${c.billCode}";
        String match1 = JxlsPlaceHolderUtils.convertPlaceHolder(pattern1);
        Assert.assertEquals("c.billCode", match1);
        String pattern2 = "${translateUtil.getConstantName('InvoiceState', c.invoiceState)}";
        String match2 = JxlsPlaceHolderUtils.convertPlaceHolder(pattern2);
        Assert.assertEquals("translateUtil.getConstantName('InvoiceState', c.invoiceState)", match2);
        String pattern3 = "￥${c.amount}";
        String match3 = JxlsPlaceHolderUtils.convertPlaceHolder(pattern3);
        Assert.assertEquals("c.amount", match3);
        String pattern4 = "${c.amount}%";
        String match4 = JxlsPlaceHolderUtils.convertPlaceHolder(pattern4);
        Assert.assertEquals("c.amount", match4);
    }

    @Test
    public void testConvertFooterCell() {
        Matcher matcher1 = JxlsPlaceHolderUtils.FOOTER_PLACE_HOLDER.matcher("$[sum(B2)]");
        Assert.assertTrue(matcher1.matches());
        String result1 = "SUM(" + matcher1.group(1) + matcher1.group(2) + ":" + matcher1.group(1) + (Integer.valueOf(matcher1.group(2)) + 10) + ")";
        Assert.assertEquals("SUM(B2:B12)", result1);
        Matcher matcher2 = JxlsPlaceHolderUtils.FOOTER_PLACE_HOLDER.matcher("合计：");
        Assert.assertFalse(matcher2.matches());
    }

    @Test
    public void testGetCellFieldName() {
        String pattern1 = "${c.billCode}";
        String match1 = JxlsPlaceHolderUtils.getCellFieldName(pattern1);
        Assert.assertEquals("billCode", match1);
        String pattern2 = "${translateUtil.getConstantName('InvoiceState', c.invoiceState)}";
        String match2 = JxlsPlaceHolderUtils.getCellFieldName(pattern2);
        Assert.assertEquals("invoiceState", match2);
        String pattern3 = "123";
        String match3 = JxlsPlaceHolderUtils.getCellFieldName(pattern3);
        Assert.assertNull(match3);
        String pattern4 = "${c.vCode}";
        String match4 = JxlsPlaceHolderUtils.getCellFieldName(pattern4);
        Assert.assertEquals("vCode", match4);
        String pattern5 = "${dateUtil.convertToDate(c.invoiceState, 'yyyyMMdd')}";
        String match5= JxlsPlaceHolderUtils.getCellFieldName(pattern5);
        Assert.assertEquals("invoiceState", match5);
    }

    @Test
    public void testConvertDateFormate() {
        String dateTimeStr = "2020-12-12 12:11:10";
        String cellValue1 = "${c.billCode}";
        String formatStr1 = JxlsPlaceHolderUtils.convertDateFormat(cellValue1, dateTimeStr);
        Assert.assertEquals("2020-12-12 12:11:10", formatStr1);
        String cellValue2 = "${dateUtil.convertToDate(c.invoiceState, 'yyyyMMdd')}";
        String formatStr2 = JxlsPlaceHolderUtils.convertDateFormat(cellValue2, dateTimeStr);
        Assert.assertEquals("20201212", formatStr2);
        String cellValue3 = "${dateUtil.convertToDate(c.invoiceState, 'yyyyMMdd HHmmss')}";
        String formatStr3 = JxlsPlaceHolderUtils.convertDateFormat(cellValue3, dateTimeStr);
        Assert.assertEquals("20201212 121110", formatStr3);
        String cellValue4 = "${dateUtil.convertToDate(c.invoiceState, 'HHmmss')}";
        String formatStr4 = JxlsPlaceHolderUtils.convertDateFormat(cellValue4, dateTimeStr);
        Assert.assertEquals("121110", formatStr4);
    }

    @Test
    public void testConvertConstantEnums() {
        String constantVal = "0";
        String cellValue1 = "${c.billCode}";
        String formatStr1 = JxlsPlaceHolderUtils.convertConstantEnums(cellValue1, constantVal);
        Assert.assertEquals("0", formatStr1);
        String cellValue2 = "${translateUtil.getConstantName('InvoiceType', c.invoiceType)}";
        String formatStr2 = JxlsPlaceHolderUtils.convertConstantEnums(cellValue2, constantVal);
        Assert.assertEquals("这是个值为0的枚举", formatStr2);
        String cellValue3 = "${translateUtil.getConstantName('AnotherField', c.anotherField)}";
        String formatStr3 = JxlsPlaceHolderUtils.convertConstantEnums(cellValue3, "1");
        Assert.assertEquals("这是个值为1的枚举", formatStr3);
    }
}
