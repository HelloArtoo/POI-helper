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
 * ����excel�����Ĵ����������
 * �Թ�������ʽ����
 * @author hurong
 */
public class ExcelUtil {
		
		/**
		 * ����һ��07 excel
		 * �˷���ֻ֧��һ��sheet
		 * 07���ϵ����ݸ�ʽ��֧�ִ����ݵ����ұ����ڴ����
		 * @param sheetName	sheet������	
		 * @param titles	������ÿ�ж�Ӧ������	������Ϊ�գ�
		 * @param cellWidth ÿ�ж�Ӧ�ĵ�Ԫ����	������Ϊ�գ�	�� 2000 , 6000, 12000... 
		 * @param dataArgNames ��data��Ӧ����ʾmap�����ݶ�Ӧ����һ�н��г���(��map������Ҫ��Ӧ	������Ϊ�գ�
		 * @param data	����Դ������	������Ϊ�գ�
		 * @throws Exception
		 */
		public static SXSSFWorkbook exportStandard07ExcelFromCollection(String sheetName, String[] titlesArr, int[] cellWidthArr, String[] dataArgNamesArr, List<Map<String,Object>> data)  throws Exception{
			
			if(titlesArr==null || cellWidthArr==null || dataArgNamesArr==null || titlesArr.length!=cellWidthArr.length || titlesArr.length!=dataArgNamesArr.length){
				throw new Exception("������С����ȷ���߹ؼ�����Ϊ�գ����⡢��Ԫ��ȡ����ݼ��ϲ����������ݳ��ȱ���һ�£�");
			}
			
			Row row = null;
			Cell cell = null;
			int index = 0;
			
			//100Ϊ�������ڴ���������ÿ100��ˢϴһ�Σ������ڴ����
			SXSSFWorkbook wb = new SXSSFWorkbook(100); 
			Sheet sheet = wb.createSheet(sheetName==null?"sheet1":sheetName);
			
			//����ÿ����Ԫ�еĿ��
			if(cellWidthArr!=null && cellWidthArr.length>0){
				for (int i = 0; i<cellWidthArr.length; i++) {
					sheet.setColumnWidth(i, cellWidthArr[i]);
				}
			}
			
			//���ñ��������̶���ʽ�����򵥵ı�����ʽ
			CellStyle style = wb.createCellStyle();
			
			row = sheet.createRow(index++);
			for (int i = 0; i<titlesArr.length; i++){
				cell = row.createCell(i);
				cell.setCellStyle(ExcelStyle.set07HeadStyle(wb,style));
				cell.setCellValue(titlesArr[i]);
			}
			
			//Excel ����
			for (Map<String,Object> map : data) {
				//����˳��map��ĳ��������excel��Ԫ���һһ��Ӧ
				row = sheet.createRow(index++);
				for (int i = 0; i<dataArgNamesArr.length; i++){
					cell = row.createCell(i);
					cell.setCellValue(map.get(dataArgNamesArr[i])==null?"":String.valueOf(map.get(dataArgNamesArr[i])));
				}
			}
			
			return wb;
		}
		
		/**
		 * ����һ��03 excel
		 * �˷���ֻ֧��һ��sheet
		 * 03���ݸ�ʽ
		 * @param sheetName	sheet������	
		 * @param titles	������ÿ�ж�Ӧ������	������Ϊ�գ�
		 * @param cellWidth ÿ�ж�Ӧ�ĵ�Ԫ����	������Ϊ�գ�	�� 2000 , 6000, 12000... 
		 * @param dataArgNames ��data��Ӧ����ʾmap�����ݶ�Ӧ����һ�н��г���(��map������Ҫ��Ӧ	������Ϊ�գ�
		 * @param data	����Դ������	������Ϊ�գ�
		 * @throws Exception
		 */
		public static HSSFWorkbook exportStandard03ExcelFromCollection(String sheetName, String[] titlesArr, int[] cellWidthArr, String[] dataArgNamesArr, List<Map<String,Object>> data)  throws Exception{
			
			if(titlesArr==null || cellWidthArr==null || dataArgNamesArr==null || titlesArr.length!=cellWidthArr.length || titlesArr.length!=dataArgNamesArr.length){
				throw new Exception("������С����ȷ���߹ؼ�����Ϊ�գ����⡢��Ԫ��ȡ����ݼ��ϲ����������ݳ��ȱ���һ�£�");
			}
			
			Row row = null;
			Cell cell = null;
			int index = 0;
			
			HSSFWorkbook wb = new HSSFWorkbook(); 
			HSSFSheet sheet = wb.createSheet(sheetName==null?"sheet1":sheetName);
			
			//����ÿ����Ԫ�еĿ��
			if(cellWidthArr!=null && cellWidthArr.length>0){
				for (int i = 0; i<cellWidthArr.length; i++) {
					sheet.setColumnWidth(i, cellWidthArr[i]);
				}
			}
			// ����һ����ʽ
			HSSFCellStyle headStyle = wb.createCellStyle();
			
			row = sheet.createRow(index++);
			for (int i = 0; i<titlesArr.length; i++){
				cell = row.createCell(i);
				cell.setCellStyle(ExcelStyle.setHeadStyle(wb, headStyle));
				cell.setCellValue(titlesArr[i]);
			}
			
			//Excel ����
			for (Map<String,Object> map : data) {
				//����˳��map��ĳ��������excel��Ԫ���һһ��Ӧ
				row = sheet.createRow(index++);
				for (int i = 0; i<dataArgNamesArr.length; i++){
					cell = row.createCell(i);
					cell.setCellValue(map.get(dataArgNamesArr[i])==null?"":map.get(dataArgNamesArr[i]).toString());
				}
			}
			
			return wb;
		}
		
		/**
		 * ��ȡ03����07��excel���ݣ�����list����
		 * �ʣ�XSSFWorkbookò�Ʊ�HSSFWorkbook���ܶ�
		 * @param file	excel�ļ�
		 * @param dataArgNamesArr	��������ÿ����map�е�Ӣ�ı�ʾ
		 * @return
		 * @throws Exception
		 */
		public static List<Map<String,Object>> readExcel4Collection(File file, String[] dataArgNamesArr) throws Exception{
			
			if(!file.exists() || dataArgNamesArr==null || dataArgNamesArr.length<=0){
				throw new Exception("����Ĳ������ԣ��ļ������ڻ���ÿ�еĲ�����Ϊ�գ�"); 
			}
			
			//�ж���xls����xlsx��excel,�Ȼ�ȡ��׺
			String suffix = file.getPath().substring( file.getPath().lastIndexOf(".")+1);
			if(!"xls".equals(suffix) && !"xlsx".equals(suffix)){
				throw new Exception("�ļ���ʽ����excel���ļ���ʽ��"); 
			}
			
			FileInputStream fs = new FileInputStream(file);
			// �õ�������
            Iterator<Row> rows = null;
			
			if("xlsx".equals(suffix)){
				//�ɶ�ȡ��������
				XSSFWorkbook wb = new XSSFWorkbook(fs);
				//��һ��sheet
				XSSFSheet sheet = wb.getSheetAt(0);
				rows = sheet.rowIterator();
			}else{
				//03��xls�Ķ���
				HSSFWorkbook wb = new HSSFWorkbook(fs);
				//��һ��sheet
				HSSFSheet sheet = wb.getSheetAt(0);
				rows = sheet.rowIterator();
			}
			
            List<Map<String,Object>> result = new ArrayList<Map<String,Object>>();
            Map<String,Object> map;
            Row row;
            Cell cell ;
            int i = 0;
            
            //ȥ����һ�б�����
            rows.next();
            
            //��������
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
