package com.gobue.excel.internal.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.gobue.excel.internal.annotation.ExcelAnnotation;
import com.gobue.excel.internal.style.ExcelStyle;

/**
 * pojo对象的形式导入导出excel 注意pojo所有的get set方法必须是get set开头
 * 比如boolean类型可能的get方法可能是isXxx 请手动改成getXxx 需要实例化对象进行操作
 * 注意vo的属性必须要注解导出的数据，例：@ExcelAnnotation(exportName="性别")
 * 
 * @author hurong
 * @param <T>
 */
@Deprecated
public class ExcelPojoUtil<T> {

    Class<T> clazz;
    // 格式化日期
    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

    public ExcelPojoUtil(Class<T> clazz) {

        this.clazz = clazz;
    }

    /**
     * exportExcel方法-poi Excel导出. for 07Excel 后缀 xlsx
     * 
     * @param title
     *            工作簿名称
     * @param dataset
     *            导出的数据集
     * @param out
     *            输出流
     */
    @SuppressWarnings("unchecked")
    public void export03Excel(String title, Collection<T> dataset, OutputStream out) throws Exception {

        // 声明一个工作薄
        // 首先检查数据看是否是正确的
        Iterator<T> iterator = dataset.iterator();
        if (dataset == null || !iterator.hasNext() || title == null || out == null) {
            throw new Exception("传入的数据参数不对！");
        }
        // 取得实际泛型类
        T tObject = iterator.next();
        Class<T> clazz = (Class<T>)tObject.getClass();

        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为20个字节
        sheet.setDefaultColumnWidth(20);
        // 生成一个样式
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置标题样式
        style = ExcelStyle.setHeadStyle(workbook, style);
        // 得到所有字段
        Field filed[] = tObject.getClass().getDeclaredFields();

        // 标题
        List<String> exportfieldtile = new ArrayList<String>();
        // 导出的字段的get方法
        List<Method> methodObj = new ArrayList<Method>();

        // 遍历整个filed
        for (int i = 0; i < filed.length; i++) {
            Field field = filed[i];
            ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
            // 如果设置了annottion
            if (excelAnnotation != null) {
                String exprot = excelAnnotation.exportName();
                // 添加到标题
                exportfieldtile.add(exprot);
                // 添加到需要导出的字段的方法
                String fieldname = field.getName();
                String getMethodName = "get" + fieldname.substring(0, 1).toUpperCase() + fieldname.substring(1);
                Method getMethod = clazz.getMethod(getMethodName, new Class[] {});
                methodObj.add(getMethod);
            }
        }

        // 产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < exportfieldtile.size(); i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(exportfieldtile.get(i));
            cell.setCellValue(text);
        }

        // 循环整个集合
        int index = 0;
        iterator = dataset.iterator();
        while (iterator.hasNext()) {
            // 从第二行开始写，第一行是标题
            index++;
            row = sheet.createRow(index);
            T t = (T)iterator.next();
            for (int k = 0; k < methodObj.size(); k++) {
                HSSFCell cell = row.createCell(k);
                Method getMethod = methodObj.get(k);
                Object value = getMethod.invoke(t, new Object[] {});
                String textValue = getValue(value);
                cell.setCellValue(textValue);
            }
        }
        workbook.write(out);
    }

    /**
     * exportExcel方法-poi Excel导出. 大数量导出，采用了SXSSFWorkbook避免内存溢出
     * 但是这个会产生非常大数量的临时文件，wb. for 07Excel 后缀 xlsx
     * 
     * @param title
     *            工作簿名称
     * @param dataset
     *            导出的数据集
     * @param out
     *            输出流
     */
    @SuppressWarnings("unchecked")
    public void export07Excel(String title, Collection<T> dataset, OutputStream out) throws Exception {

        // 声明一个工作薄
        // 首先检查数据看是否是正确的
        Iterator<T> iterator = dataset.iterator();
        if (dataset == null || !iterator.hasNext() || title == null || out == null) {
            throw new Exception("传入的数据参数不对！");
        }
        // 取得实际泛型类
        T tObject = iterator.next();
        Class<T> clazz = (Class<T>)tObject.getClass();

        // 100为保存在内存中行数，每100行刷洗一次，避免内存溢出
        SXSSFWorkbook workbook = new SXSSFWorkbook(100);
        Sheet sheet = workbook.createSheet(title == null ? "sheet1" : title);
        // 设置表格默认列宽度为20个字节
        sheet.setDefaultColumnWidth(20);
        // 生成一个样式
        CellStyle style = workbook.createCellStyle();
        // 设置标题样式
        style = ExcelStyle.set07HeadStyle(workbook, style);
        // 得到所有字段
        Field filed[] = tObject.getClass().getDeclaredFields();

        // 标题
        List<String> exportfieldtile = new ArrayList<String>();
        // 导出的字段的get方法
        List<Method> methodObj = new ArrayList<Method>();

        // 遍历整个filed
        for (int i = 0; i < filed.length; i++) {
            Field field = filed[i];
            ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
            // 如果设置了annottion
            if (excelAnnotation != null) {
                String exprot = excelAnnotation.exportName();
                // 添加到标题
                exportfieldtile.add(exprot);
                // 添加到需要导出的字段的方法
                String fieldname = field.getName();
                String getMethodName = "get" + fieldname.substring(0, 1).toUpperCase() + fieldname.substring(1);
                Method getMethod = clazz.getMethod(getMethodName, new Class[] {});
                methodObj.add(getMethod);
            }
        }

        // 产生表格标题行
        Row row = sheet.createRow(0);
        for (int i = 0; i < exportfieldtile.size(); i++) {
            Cell cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(exportfieldtile.get(i));
            cell.setCellValue(text);
        }

        // 循环整个集合
        int index = 0;
        iterator = dataset.iterator();
        while (iterator.hasNext()) {
            // 从第二行开始写，第一行是标题
            index++;
            row = sheet.createRow(index);
            T t = (T)iterator.next();
            for (int k = 0; k < methodObj.size(); k++) {
                Cell cell = row.createCell(k);
                Method getMethod = methodObj.get(k);
                Object value = getMethod.invoke(t, new Object[] {});
                String textValue = getValue(value);
                cell.setCellValue(textValue);
            }
        }
        workbook.write(out);
        // 临时文件的处理
        workbook.dispose();
    }

    /**
     * 返回集合poi导入
     * 
     * @param file
     *            导入的文件
     * @param pattern
     * @return
     */
    public Collection<T> importExcel(File file, String... pattern) throws Exception {

        // 判断是03还是07格式的excel
        String suffix = file.getPath().substring(file.getPath().lastIndexOf(".") + 1);
        if (!"xls".equals(suffix) && !"xlsx".equals(suffix)) {
            throw new Exception("文件格式不是excel的文件格式！");
        }
        Collection<T> dist = new ArrayList<T>();
        /**
         * 类反射得到调用方法
         */
        // 得到目标目标类的所有的字段列表
        Field[] fields = clazz.getDeclaredFields();
        // 将所有标有Annotation的字段，也就是允许导入数据的字段,放入到一个map中
        Map<String, Method> fieldMap = new HashMap<String, Method>();
        // 循环读取所有字段
        for (Field field : fields) {
            // 得到单个字段上的Annotation
            ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
            // 如果标识了Annotationd
            if (excelAnnotation != null) {
                String fieldName = field.getName();
                // 构造设置了Annotation的字段的Setter方法
                String setMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                // 构造调用的method
                Method setMethod = clazz.getMethod(setMethodName, new Class[] { field.getType() });
                // 将这个method以Annotaion的名字为key来存入
                fieldMap.put(excelAnnotation.exportName(), setMethod);
            }
        }

        Iterator<Row> row = null;

        /**
         * excel的解析开始
         */
        // 将传入的File构造为FileInputStream;
        FileInputStream inputStream = new FileInputStream(file);
        if ("xlsx".equals(suffix)) {
            // 可读取大数据量
            XSSFWorkbook wb = new XSSFWorkbook(inputStream);
            // 第一个sheet
            XSSFSheet sheet = wb.getSheetAt(0);
            row = sheet.rowIterator();
        } else {
            // 得到工作表
            HSSFWorkbook book = new HSSFWorkbook(inputStream);
            // 得到第一页
            HSSFSheet sheet = book.getSheetAt(0);
            // 得到第一面的所有行
            row = sheet.rowIterator();
        }

        /**
         * 标题解析
         */
        // 得到第一行，也就是标题行
        Row titleRow = row.next();
        // 得到第一行的所有列
        Iterator<Cell> cellTitle = titleRow.cellIterator();
        // 将标题的文字内容放入到一个map中
        Map<Integer, String> titleMap = new HashMap<Integer, String>();
        // 从标题第一列开始
        int i = 0;
        // 循环标题所有的列
        while (cellTitle.hasNext()) {
            Cell cell = (Cell)cellTitle.next();
            String value = cell.getStringCellValue();
            titleMap.put(i, value);
            i++;
        }

        /**
         * 解析内容行
         */
        while (row.hasNext()) {
            // 标题下的第一行
            Row rown = row.next();
            // 行的所有列
            Iterator<Cell> cellBody = rown.cellIterator();
            // 得到传入类的实例
            T tObject = clazz.newInstance();
            // 遍历一行的列
            int col = 0;
            while (cellBody.hasNext()) {
                Cell cell = (Cell)cellBody.next();
                // 这里得到此列的对应的标题
                String titleString = titleMap.get(col++);
                // 如果这一列的标题和类中的某一列的Annotation相同，那么则调用此类的的set方法，进行设值
                if (fieldMap.containsKey(titleString)) {
                    Method setMethod = fieldMap.get(titleString);
                    // 得到setter方法的参数
                    Type[] types = setMethod.getGenericParameterTypes();
                    // 只要一个参数
                    String xclass = String.valueOf(types[0]);
                    // 判断参数类型
                    if ("class java.lang.String".equals(xclass)) {
                        setMethod.invoke(tObject, cell.getStringCellValue());
                    } else if ("class java.util.Date".equals(xclass)) {
                        // setMethod.invoke(tObject,
                        // cell.getDateCellValue());
                        setMethod.invoke(tObject, sdf.parse(cell.getStringCellValue()));
                    } else if ("class java.lang.Boolean".equals(xclass) || "boolean".equals(xclass)) {
                        Boolean boolName = true;
                        if ("否".equals(cell.getStringCellValue())) {
                            boolName = false;
                        }
                        setMethod.invoke(tObject, boolName);
                    } else if ("class java.lang.Integer".equals(xclass) || "int".equals(xclass)) { // 基本类型
                                                                                                   // 和
                                                                                                   // 包装类型
                        // setMethod.invoke(tObject, new
                        // Integer(String.valueOf((int)cell.getNumericCellValue())));
                        setMethod.invoke(tObject, new Integer(cell.getStringCellValue()));
                    } else if ("class java.lang.Long".equals(xclass) || "long".equals(xclass)) {
                        setMethod.invoke(tObject, new Long(cell.getStringCellValue()));
                    } else if ("class java.lang.Double".equals(xclass) || "double".equals(xclass)) {
                        setMethod.invoke(tObject, new Double(cell.getStringCellValue()));
                    } else {
                        // 需要别的类型，这里补充
                    }
                }
            }
            dist.add(tObject);
        }
        return dist;
    }

    /**
     * getValue方法-cell值处理.
     * 
     * @param value
     * @return
     */
    public String getValue(Object value) {

        String textValue = "";
        if (value == null) {
            return textValue;
        }
        if (value instanceof Boolean) {
            boolean bValue = (Boolean)value;
            textValue = "是";
            if (!bValue) {
                textValue = "否";
            }
        } else if (value instanceof Date) {
            Date date = (Date)value;
            textValue = sdf.format(date);

        } else {
            textValue = value.toString();
        }
        return textValue;
    }

}
