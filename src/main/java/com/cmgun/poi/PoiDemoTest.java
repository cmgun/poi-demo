package com.cmgun.poi;

import java.util.ArrayList;
import java.util.List;

/**
 * 使用ali-easyexcel
 * https://github.com/alibaba/easyexcel
 */
public class PoiDemoTest {

    public static void main(String[] args) {
        System.out.println(System.getProperty("java.io.tmpdir"));
        // 根据模板写入excel
        System.out.println("500rows * 5cols data prepare...");
        List<Entity> data1 = createData(500);
        long startTime = System.currentTimeMillis();
        System.out.println("500rows * 5cols start export, using template...");
        PoiUtil.export("template1.xlsx", "test1.xlsx", data1, 1);
        System.out.println("500rows * 5cols, 耗时:" + (System.currentTimeMillis() - startTime));

        // 根据javaEntity的注解表头写入
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

    public static List<Entity> createData(int length) {
        List<Entity> data = new ArrayList<>();
        for (int i = 1; i <= length; i++) {
            data.add(new Entity(i, "msg" + i));
        }
        return data;
    }

    /**
     * 注解模板测试
     */
    public static void testAnnotationTemplate() {

    }
}
