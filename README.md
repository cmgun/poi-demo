# poi-demo
EasyExcel增强处理。增强功能如下：
- 支持处理jxls-poi-jdk1.6的模板解析的导出。

## jxls-poi-jdk1.6的模板解析的导出设计流程
1. 解析excel模板中的最后一行，作为cell列模板。
2. 解析列模板，读取其中带有 c.* 的内容，* 为需要反射读取的bean的字段，反射字段记录到反射字段列表中，与cell列模板一一对应。
3. 写数据，根据反射列表读取实际属性的值，判断cell列模板中是否有工具类引入。目前只支持处理 枚举类转换 和 日期格式化 。

具体支持操作见示例：[地址](https://github.com/cmgun/poi-demo/blob/master/src/main/java/com/cmgun/poi/PoiDemo.java)

增强 EasyExcelFactory，新增支持上述操作的 Writer 和配套的 ExcelBuilder。
