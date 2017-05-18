# poi-helper

A Java library for reading and writing Microsoft Office Excel based on Apache POI.

## Features

- 链式编程，简单实用，丰富的Excel导入导出功能。
- 格式丰富，支持2003版Excel(.xls),和2007版以后的Excel(.xlsx)。
- 支持非严格的动态列数据导入导出。Import数据存储在Map中，Export需要指定List的数据和动态header。
- 支持手动追加数据列(针对读取)。场景如Excel数据格式错误时候，错误列追加到最后写入到新的文件。
- 支持指定Sheet读取，默认从0个开始。
- 支持最大导入导出行数限制，动态设置。
- 支持校验数据表头，动态设置。
- 支持Excel公式数据正确读取，目前默认读取为公式计算的结果。
- 支持动态属性设置用户自定义参数关系。比如将动态列【北京】替换成系统可读的【6】。
- 支持导出文件时设置统一的必填项样式。

## Getting started

### Read
```java
ExcelUtility<TestExcelModel> utility = ExcelUtility
    .newImportBuilder(new File(path), TestExcelModel.class, properties).maxUploadSum(2000).build();
List<TestExcelModel> fetchData = utility.readExcel();
```
### Read dynamic columns
```java
Map<String, Object> parameters = new HashMap<String, Object>();
parameters.put("北京", 6);// 北京的值改成6
parameters.put("上海", 10);// 上海的值改成10
parameters.put("广州", "不知道");// 广州的值改成10

ExcelUtility<TestExcelModel> utility = ExcelUtility
    .newImportBuilder(new File(dynamicPath), TestExcelModel.class, properties).maxUploadSum(2000)
    .dynamicProperty("dcQty").dynamicColumnParameters(parameters).build();

List<TestExcelModel> fetchData = utility.readExcel();
```

### Write
```java
ExcelUtility<TestExcelModel> utility = ExcelUtility
    .newExportBuilder(fullPathName, expectedHeaders, properties, contentList).maxExportNum(1000).build();
File writeExcel = utility.writeExcel();
```

### Write dynamic columns
```java
ExcelUtility<TestExcelModel> build = ExcelUtility
    .newExportBuilder(fullPathName, expectedHeaders, properties, contentList)
    .dynamicProperty4List("dcQty4Export", new String[] { "北京", "上海", "广州" }).build();
File writeExcel = build.writeExcel();
```

### Append msg column
```java
ExcelUtility<TestExcelModel> utility = ExcelUtility.newImportBuilder(new File(path), TestExcelModel.class,
    properties).build();

List<TestExcelModel> fetchData = utility.readExcel();

// 读取完了。发现有错误
Map<Integer, String> msgMap = new HashMap<Integer, String>(fetchData.size(), 1);
for (int i = 1; i < fetchData.size(); i++) {
    if (i % 2 == 0) {
        msgMap.put(i, "错误，这是个错误。");
    }
}

// 追加错误列
File file = utility.appendExcelColumn("详细失败信息", msgMap, EXCEL_PREFIX);
```

## Others

 Improving...