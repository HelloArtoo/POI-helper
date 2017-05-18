package com.gobue.excel.api;

import static org.hamcrest.core.Is.is;
import static org.junit.Assert.assertThat;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.junit.Test;

import com.gobue.excel.fixture.TestExcelModel;
import com.gobue.excel.internal.exception.POIException;

/**
 * @author hurong
 * 
 * @description: 用例测试
 */
public class ExcelUtilityTest {

    private final String[] expectedHeaders = { "商品编号*", "商品名称*", "DC*", "是否补货*", "库存满足率*", "BAND*", "公式列*", "日期*" };

    private final String[] properties = { "sku", "productName", "dc", "isRep", "cr", "band", "formulaRst", "createDate" };

    private final String path = ExcelUtilityTest.class.getResource("/excel/test_excel_2007.xlsx").getFile();

    private final String dynamicPath = ExcelUtilityTest.class.getResource("/excel/test_excel_2007_dynamic.xlsx")
        .getFile();

    private final int sheet = 0; // 第一个

    private final String fullPathName = "D://utility//test_export.xls";

    private final String EXCEL_PREFIX = "上传更新错误信息";

    private List<TestExcelModel> contentList = new ArrayList<TestExcelModel>();

    {
        for (int i = 0; i < 100; i++) {
            TestExcelModel model = new TestExcelModel();
            model.setSku(100000L + i);
            model.setProductName("苹果手机Iphone " + i);
            model.setIsRep(i % 2 == 0 ? true : false);
            model.setFormulaRst(i / 3D);
            model.setDc(i);
            model.setCreateDate(new Date());
            model.setCr(i % 2 == 0 ? 0.5 : 0.2);
            model.setBand(i % 3 == 0 ? 'A' : 'E');

            List<Integer> qtys = new ArrayList<Integer>();
            for (int j = 0; j < 3; j++) {
                qtys.add(i * j * 3);
            }
            model.setDcQty4Export(qtys);
            contentList.add(model);
        }
    }

    @Test
    public void assertFileExist() throws IOException {

        File file = new File(dynamicPath);
        System.out.println("file.getAbsolutePath():" + file.getAbsolutePath());
        System.out.println("file.getCanonicalPath():" + file.getCanonicalPath());
        System.out.println("file.getPath():" + file.getPath());
        String path = file.getCanonicalPath();
        System.out.println(path.substring(0, path.lastIndexOf('\\') + 1));
        assertThat(file.exists(), is(true));
    }

    /**
     * 测试读取
     * 
     * @throws IOException
     * @throws POIException
     * @throws SecurityException
     * @throws IllegalArgumentException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws ClassNotFoundException
     * @throws NoSuchFieldException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     */
    @Test
    public void assertReadExcel() throws IOException, POIException, SecurityException, IllegalArgumentException,
                                 InstantiationException, IllegalAccessException, ClassNotFoundException,
                                 NoSuchFieldException, NoSuchMethodException, InvocationTargetException {

        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newImportBuilder(new File(path), TestExcelModel.class, properties).maxUploadSum(2000).build();

        List<TestExcelModel> fetchData = utility.readExcel();
        assertThat(fetchData.size(), is(9));
    }

    /**
     * 测试写入
     * 
     * @throws POIException
     * @throws IOException
     * @throws SecurityException
     * @throws IllegalArgumentException
     * @throws NoSuchFieldException
     * @throws NoSuchMethodException
     * @throws IllegalAccessException
     * @throws InvocationTargetException
     */
    @Test
    public void assertWriteExcel() throws POIException, IOException, SecurityException, IllegalArgumentException,
                                  NoSuchFieldException, NoSuchMethodException, IllegalAccessException,
                                  InvocationTargetException {

        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newExportBuilder(fullPathName, expectedHeaders, properties, contentList).maxExportNum(1000).build();
        utility.writeExcel();
        assertThat(new File(fullPathName).exists(), is(true));
    }

    // 测试必填项导出
    @Test
    public void assertWriteRequiredColumnExcel() throws POIException, IOException, SecurityException,
                                                IllegalArgumentException, NoSuchFieldException, NoSuchMethodException,
                                                IllegalAccessException, InvocationTargetException {

        // expectedHeaders中的 "商品编号*", "DC*"为必填项
        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newExportBuilder(fullPathName, expectedHeaders, properties, contentList).requiredColumn("商品编号*")
            .requiredColumn("DC*").maxExportNum(1000).build();
        utility.writeExcel();
        assertThat(new File(fullPathName).exists(), is(true));
    }

    // 动态列读取
    @Test
    public void assertReadDynamicColumnExcel() throws IOException, POIException, SecurityException,
                                              IllegalArgumentException, InstantiationException, IllegalAccessException,
                                              ClassNotFoundException, NoSuchFieldException, NoSuchMethodException,
                                              InvocationTargetException {

        Map<String, Object> parameters = new HashMap<String, Object>();
        parameters.put("北京", 6);// 北京的值改成6
        parameters.put("上海", 10);// 上海的值改成10
        parameters.put("广州", "不知道");// 广州的值改成10

        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newImportBuilder(new File(dynamicPath), TestExcelModel.class, properties).maxUploadSum(2000)
            .dynamicProperty("dcQty").dynamicColumnParameters(parameters).build();

        List<TestExcelModel> fetchData = utility.readExcel();
        System.out.println(Arrays.toString(utility.getHeaders()));
        System.out.println(fetchData.get(0).getDcQty());
        assertThat(fetchData.size(), is(9));
    }

    // 动态列写入
    @Test
    public void assertWriteDynamicColumnExcel() throws POIException, IOException, SecurityException,
                                               IllegalArgumentException, NoSuchFieldException, NoSuchMethodException,
                                               IllegalAccessException, InvocationTargetException {

        ExcelUtility<TestExcelModel> build = ExcelUtility
            .newExportBuilder(fullPathName, expectedHeaders, properties, contentList)
            .dynamicProperty4List("dcQty4Export", new String[] { "北京", "上海", "广州" }).build();
        build.writeExcel();
        assertThat(new File(fullPathName).exists(), is(true));
    }

    /**
     * 写入错误信息列，返回文件供下载
     * 
     * @throws IOException
     * @throws POIException
     * @throws SecurityException
     * @throws IllegalArgumentException
     * @throws InstantiationException
     * @throws IllegalAccessException
     * @throws ClassNotFoundException
     * @throws NoSuchFieldException
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     */
    @Test
    public void assertReadAndAppendExcelColumn() throws IOException, POIException, SecurityException,
                                                IllegalArgumentException, InstantiationException,
                                                IllegalAccessException, ClassNotFoundException, NoSuchFieldException,
                                                NoSuchMethodException, InvocationTargetException {

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
        System.out.println(file.getCanonicalPath());
        assertThat(file.exists(), is(true));
    }

    @Test
    public void assertWrongOperation() throws POIException, IOException, SecurityException, IllegalArgumentException,
                                      InstantiationException, IllegalAccessException, ClassNotFoundException,
                                      NoSuchFieldException, NoSuchMethodException, InvocationTargetException {

        ExcelUtility<TestExcelModel> exp = ExcelUtility.newExportBuilder(fullPathName, expectedHeaders, properties,
            contentList).build();
        ExcelUtility<TestExcelModel> imp = ExcelUtility.newImportBuilder(new File(path), TestExcelModel.class,
            properties).build();
        imp.writeExcel();
        assertThat(exp.readExcel().size(), is(0));
    }

    @Test(expected = POIException.class)
    public void assertWrongHeaders() throws POIException, IOException {

        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newImportBuilder(new File(path), TestExcelModel.class, properties)
            .headers(new String[] { "wrong headers" }).build();
    }

    @Test(expected = POIException.class)
    public void assertNegativeSheetNo() throws POIException, IOException {

        ExcelUtility<TestExcelModel> utility = ExcelUtility
            .newImportBuilder(new File(path), TestExcelModel.class, properties).sheet(-1).build();
    }

}
