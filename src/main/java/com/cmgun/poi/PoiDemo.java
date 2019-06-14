package com.cmgun.poi;

import com.cmgun.util.DateUtil;
import com.cmgun.util.TranslateUtil;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 使用ali-easyexcel
 * https://github.com/alibaba/easyexcel
 */
public class PoiDemo {

    public static void main(String[] args) {
        System.out.println(System.getProperty("java.io.tmpdir"));
        // 根据Jxls模板写入excel
        testExcelTemplate1();
//        testExcelTemplate2();

        // 根据javaEntity的注解表头写入
//        testAnnotationTemplate();
        // 模板样式
//        testEasyExcelTemplate();
    }

    private static List<Entity> createData(int length) {
        List<Entity> data = new ArrayList<>();
        for (int i = 1; i <= length; i++) {
            data.add(new Entity(i, "msg" + i, i % 2 == 0 ? "0" : "1"));
        }
        return data;
    }

    private static List<JxlsEntity> createJxlsDasta(int size) {
        List<JxlsEntity> data = new ArrayList<>();
        for (int i = 1; i <= size; i++) {
            data.add(new JxlsEntity(i, "msg" + i, i % 2 == 0 ? "0" : "1"));
        }
        return data;
    }

    private static Map<String, Object> createJxlsTmpDatas(int size) {
        Map<String, Object> datas = new HashMap<>();
        List<JxlsEntity> data = new ArrayList<>();
        for (int i = 1; i <= size; i++) {
            data.add(new JxlsEntity(i, "msg" + i, i % 2 == 0 ? "0" : "1"));
        }
        datas.put("datas", data);
        datas.put("dateUtil", new DateUtil());
        datas.put("translateUtil", new TranslateUtil());
        return datas;
    }

    public static void testExcelTemplate1() {
        System.out.println("[jexl] 500rows * 7cols data prepare...");
//        List<JxlsEntity> data1 = createJxlsDasta(500);
        Map<String, Object> datas = createJxlsTmpDatas(500);
        long startTime = System.currentTimeMillis();
        System.out.println("[jexl] 500rows * 7cols start export, using template...");
        PoiUtil.exportForJxlsTemp("template11.xlsx", "test13.xlsx", datas);
        System.out.println("[jexl] 500rows * 7cols, 耗时:" + (System.currentTimeMillis() - startTime));
    }

    public static void testExcelTemplate2() {
        System.out.println("[non] 500rows * 7cols data prepare...");
        List<JxlsEntity> data1 = createJxlsDasta(500);
//        Map<String, Object> datas = createJxlsTmpDatas(500);
        long startTime = System.currentTimeMillis();
        System.out.println("[non] 500rows * 7cols start export, using template...");
        PoiUtil.exportForJxlsTemp1("template.xlsx", "test11.xlsx", data1);
        System.out.println("[non] 500rows * 7cols, 耗时:" + (System.currentTimeMillis() - startTime));
    }

    public static void testEasyExcelTemplate() {
        System.out.println("500rows * 5cols data prepare...");
        List<Entity> data2 = createData(500);
        long startTime1 = System.currentTimeMillis();
        System.out.println("500rows * 5cols start export...");
        PoiUtil.export("template1.xslx", "test2.xlsx", data2, 3);
        System.out.println("500rows * 5cols, 耗时:" + (System.currentTimeMillis() - startTime1));
    }

    /**
     * 注解模板测试
     */
    public static void testAnnotationTemplate() {
        System.out.println("500rows * 5cols data prepare...");
        List<Entity> data2 = createData(500);
        long startTime1 = System.currentTimeMillis();
        System.out.println("500rows * 5cols start export...");
        PoiUtil.export("test2.xlsx", data2);
        System.out.println("500rows * 5cols, 耗时:" + (System.currentTimeMillis() - startTime1));

        System.out.println("5000rows * 5cols data prepare...");
        List<Entity> data3 = createData(5000);
        long startTime2 = System.currentTimeMillis();
        System.out.println("5000rows * 5cols start export...");
        PoiUtil.export("test3.xlsx", data3);
        System.out.println("5000rows * 5cols, 耗时:" + (System.currentTimeMillis() - startTime2));

        System.out.println("50000rows * 5cols data prepare...");
        List<Entity> data4 = createData(50000);
        long startTime3 = System.currentTimeMillis();
        System.out.println("50000rows * 5cols start export...");
        PoiUtil.export("test4.xlsx", data3);
        System.out.println("50000rows * 5cols, 耗时:" + (System.currentTimeMillis() - startTime3));
    }
}
