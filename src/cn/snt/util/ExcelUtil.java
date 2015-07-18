package cn.snt.util;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 鉴于excel操作的大量冗余操作
 * 以工具类形式操作
 * @author hurong
 */
public class ExcelUtil {
		
		/**
		 * 导出一个07 excel
		 * 此方法只支持一个sheet
		 * 07以上的数据格式，支持大数据导出且避免内存溢出
		 * @param sheetName	sheet的名称	
		 * @param titles	标题栏每列对应的名称	（不能为空）
		 * @param cellWidth 每列对应的单元格宽度	（不能为空）	如 2000 , 6000, 12000... 
		 * @param dataArgNames 与data对应，表示map中数据对应在哪一列进行呈现(与map中数据要对应	（不能为空）
		 * @param data	数据源，集合	（不能为空）
		 * @throws Exception
		 */
		public static SXSSFWorkbook exportStandard07ExcelFromCollection(String sheetName, String[] titlesArr, int[] cellWidthArr, String[] dataArgNamesArr, List<Map<String,Object>> data)  throws Exception{
			
			if(titlesArr==null || cellWidthArr==null || dataArgNamesArr==null || titlesArr.length!=cellWidthArr.length || titlesArr.length!=dataArgNamesArr.length){
				throw new Exception("参数大小不正确或者关键参数为空，标题、单元宽度、数据集合参数名的数据长度必须一致！");
			}
			
			Row row = null;
			Cell cell = null;
			int index = 0;
			
			//100为保存在内存中行数，每100行刷洗一次，避免内存溢出
			SXSSFWorkbook wb = new SXSSFWorkbook(100); 
			Sheet sheet = wb.createSheet(sheetName==null?"sheet1":sheetName);
			
			//设置每个单元列的宽度
			if(cellWidthArr!=null && cellWidthArr.length>0){
				for (int i = 0; i<cellWidthArr.length; i++) {
					sheet.setColumnWidth(i, cellWidthArr[i]);
				}
			}
			
			//设置标题栏，固定样式，即简单的标题样式
			CellStyle style = wb.createCellStyle();
			
			row = sheet.createRow(index++);
			for (int i = 0; i<titlesArr.length; i++){
				cell = row.createCell(i);
				cell.setCellStyle(ExcelStyle.set07HeadStyle(wb,style));
				cell.setCellValue(titlesArr[i]);
			}
			
			//Excel 内容
			for (Map<String,Object> map : data) {
				//按照顺序，map中某个数据与excel单元格的一一对应
				row = sheet.createRow(index++);
				for (int i = 0; i<dataArgNamesArr.length; i++){
					cell = row.createCell(i);
					cell.setCellValue(map.get(dataArgNamesArr[i])==null?"":String.valueOf(map.get(dataArgNamesArr[i])));
				}
			}
			
			return wb;
		}
		
		/**
		 * 导出一个03 excel
		 * 此方法只支持一个sheet
		 * 03数据格式
		 * @param sheetName	sheet的名称	
		 * @param titles	标题栏每列对应的名称	（不能为空）
		 * @param cellWidth 每列对应的单元格宽度	（不能为空）	如 2000 , 6000, 12000... 
		 * @param dataArgNames 与data对应，表示map中数据对应在哪一列进行呈现(与map中数据要对应	（不能为空）
		 * @param data	数据源，集合	（不能为空）
		 * @throws Exception
		 */
		public static HSSFWorkbook exportStandard03ExcelFromCollection(String sheetName, String[] titlesArr, int[] cellWidthArr, String[] dataArgNamesArr, List<Map<String,Object>> data)  throws Exception{
			
			if(titlesArr==null || cellWidthArr==null || dataArgNamesArr==null || titlesArr.length!=cellWidthArr.length || titlesArr.length!=dataArgNamesArr.length){
				throw new Exception("参数大小不正确或者关键参数为空，标题、单元宽度、数据集合参数名的数据长度必须一致！");
			}
			
			Row row = null;
			Cell cell = null;
			int index = 0;
			
			HSSFWorkbook wb = new HSSFWorkbook(); 
			HSSFSheet sheet = wb.createSheet(sheetName==null?"sheet1":sheetName);
			
			//设置每个单元列的宽度
			if(cellWidthArr!=null && cellWidthArr.length>0){
				for (int i = 0; i<cellWidthArr.length; i++) {
					sheet.setColumnWidth(i, cellWidthArr[i]);
				}
			}
			// 生成一个样式
			HSSFCellStyle headStyle = wb.createCellStyle();
			
			row = sheet.createRow(index++);
			for (int i = 0; i<titlesArr.length; i++){
				cell = row.createCell(i);
				cell.setCellStyle(ExcelStyle.setHeadStyle(wb, headStyle));
				cell.setCellValue(titlesArr[i]);
			}
			
			//Excel 内容
			for (Map<String,Object> map : data) {
				//按照顺序，map中某个数据与excel单元格的一一对应
				row = sheet.createRow(index++);
				for (int i = 0; i<dataArgNamesArr.length; i++){
					cell = row.createCell(i);
					cell.setCellValue(map.get(dataArgNamesArr[i])==null?"":map.get(dataArgNamesArr[i]).toString());
				}
			}
			
			return wb;
		}
		
		/**
		 * 读取03或者07的excel数据，返回list集合
		 * 问：XSSFWorkbook貌似比HSSFWorkbook慢很多
		 * @param file	excel文件
		 * @param dataArgNamesArr	用来保存每列在map中的英文表示
		 * @return
		 * @throws Exception
		 */
		public static List<Map<String,Object>> readExcel4Collection(File file, String[] dataArgNamesArr) throws Exception{
			
			if(!file.exists() || dataArgNamesArr==null || dataArgNamesArr.length<=0){
				throw new Exception("传入的参数不对，文件不存在或者每列的参数名为空！"); 
			}
			
			//判断是xls还是xlsx的excel,先获取后缀
			String suffix = file.getPath().substring( file.getPath().lastIndexOf(".")+1);
			if(!"xls".equals(suffix) && !"xlsx".equals(suffix)){
				throw new Exception("文件格式不是excel的文件格式！"); 
			}
			
			FileInputStream fs = new FileInputStream(file);
			// 得到所有行
            Iterator<Row> rows = null;
			
			if("xlsx".equals(suffix)){
				//可读取大数据量
				XSSFWorkbook wb = new XSSFWorkbook(fs);
				//第一个sheet
				XSSFSheet sheet = wb.getSheetAt(0);
				rows = sheet.rowIterator();
			}else{
				//03的xls的读法
				HSSFWorkbook wb = new HSSFWorkbook(fs);
				//第一个sheet
				HSSFSheet sheet = wb.getSheetAt(0);
				rows = sheet.rowIterator();
			}
			
            List<Map<String,Object>> result = new ArrayList<Map<String,Object>>();
            Map<String,Object> map;
            Row row;
            Cell cell ;
            int i = 0;
            
            //去掉第一行标题行
            rows.next();
            
            //解析内容
            while (rows.hasNext()) {
            	row= rows.next();
            	map = new HashMap<String,Object>();
            	 Iterator<Cell> cellBody = row.cellIterator();
            	 while (cellBody.hasNext()) {
            		 cell = (Cell) cellBody.next();
            		 map.put(dataArgNamesArr[i++], cell.getStringCellValue()==null?null:cell.getStringCellValue().trim());
				}
            	 result.add(map);
            	 i=0;
            }
			
			return result;
		}
		
}
