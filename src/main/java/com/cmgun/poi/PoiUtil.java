package com.cmgun.poi;


import java.io.File;
import java.util.List;

public class PoiUtil {

    public static void export(String templateFileName, String targetFileName, List<?> params) {
        // get template file
        File template = new File(templateFileName);
        if (!template.exists()) {
            System.err.println("template file does not exist");
            return;
        }
    }
}
