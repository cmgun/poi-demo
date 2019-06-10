package com.cmgun.poi;

public class PoiUtil {

    public static void export(Workbook wb,List<Map<String,String>> params,String sheetName) {
        // 获取模板中的sheet
        Sheet sheet = wb.getSheetAt(0);
        // 设置模板页的名称
        wb.setSheetName(0, sheetName);
        // 获得模板航
        Row tmpRow = sheet.getRow(1);
        // 获得非空白的行数
        int last = sheet.getLastRowNum();
        // 循环遍历填充数据
        for (int i = 0, len = params.size(); i < len; i++){
            // 获的开始填充的一行 就是模板的下一行
            int index = i+last+1;
            Map<String,String> map = params.get(i);
            // 创建新的一行
            Row row = sheet.getRow(index);
            if(row == null) {
                row = sheet.createRow(index);
            }
            // 循环便利模板行的列 获取${key}中的key
            for (int j = tmpRow.getFirstCellNum() ; j < tmpRow.getLastCellNum() ; j++){
                // 得到模板行的第j列单元格
                Cell tmpCell = tmpRow.getCell(j);
                // 获取key
                String key = tmpCell.getStringCellValue().replace("$", "").replace("{", "").replace("}", "");
                int columnindex = tmpCell.getColumnIndex();
                System.out.println(MessageFormat.format("这是第{0}行，第{1}列的key：{2}",index,columnindex,key));
                // 得到创建的一行的第j列单元格
                Cell c = row.getCell(j);
                if(c == null)
                    c = row.createCell(columnindex);
                // 填充单元格数据
                c.setCellValue(map.get(key));
            }
        }
        // 删除模板行
        sheet.shiftRows(2,5,-1);
    }
}
