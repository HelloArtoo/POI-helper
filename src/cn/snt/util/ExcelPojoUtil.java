package cn.snt.util;

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


/**
 * pojo�������ʽ���뵼��excel
 * ע��pojo���е�get set����������get set��ͷ
 * ����boolean���Ϳ��ܵ�get����������isXxx ���ֶ��ĳ�getXxx
 * ��Ҫʵ����������в���
 * ע��vo�����Ա���Ҫע�⵼�������ݣ�����@ExcelAnnotation(exportName="�Ա�")
 * @author hurong
 * @param <T>
 */
public class ExcelPojoUtil<T> {

	Class<T> clazz;
	// ��ʽ������
	SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");

	public ExcelPojoUtil(Class<T> clazz) {
		this.clazz = clazz;
	}

	/**
	 * exportExcel����-poi Excel����.
	 * for 07Excel 
	 * ��׺ xlsx
	 * @param title
	 *            ����������
	 * @param dataset
	 *            ���������ݼ�
	 * @param out
	 *            �����
	 */
	@SuppressWarnings("unchecked")
	public void export03Excel(String title, Collection<T> dataset, OutputStream out) throws Exception {
		// ����һ��������
		// ���ȼ�����ݿ��Ƿ�����ȷ��
		Iterator<T> iterator = dataset.iterator();
		if (dataset == null || !iterator.hasNext() || title == null || out == null) {
			throw new Exception("��������ݲ������ԣ�");
		}
		// ȡ��ʵ�ʷ�����
		T tObject = iterator.next();
		Class<T> clazz = (Class<T>) tObject.getClass();

		HSSFWorkbook workbook = new HSSFWorkbook();
		// ����һ�����
		HSSFSheet sheet = workbook.createSheet(title);
		// ���ñ��Ĭ���п��Ϊ20���ֽ�
		sheet.setDefaultColumnWidth(20);
		// ����һ����ʽ
		HSSFCellStyle style = workbook.createCellStyle();
		// ���ñ�����ʽ
		style = ExcelStyle.setHeadStyle(workbook, style);
		// �õ������ֶ�
		Field filed[] = tObject.getClass().getDeclaredFields();

		// ����
		List<String> exportfieldtile = new ArrayList<String>();
		// �������ֶε�get����
		List<Method> methodObj = new ArrayList<Method>();
		
		// ��������filed
		for (int i = 0; i < filed.length; i++) {
			Field field = filed[i];
			ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
			// ���������annottion
			if (excelAnnotation != null) {
				String exprot = excelAnnotation.exportName();
				// ��ӵ�����
				exportfieldtile.add(exprot);
				// ��ӵ���Ҫ�������ֶεķ���
				String fieldname = field.getName();
				String getMethodName = "get" + fieldname.substring(0, 1).toUpperCase() + fieldname.substring(1);
				Method getMethod = clazz.getMethod(getMethodName, new Class[] {});
				methodObj.add(getMethod);
			}
		}

		// ������������
		HSSFRow row = sheet.createRow(0);
		for (int i = 0; i < exportfieldtile.size(); i++) {
			HSSFCell cell = row.createCell(i);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(exportfieldtile.get(i));
			cell.setCellValue(text);
		}

		// ѭ����������
		int index = 0;
		iterator = dataset.iterator();
		while (iterator.hasNext()) {
			// �ӵڶ��п�ʼд����һ���Ǳ���
			index++;
			row = sheet.createRow(index);
			T t = (T) iterator.next();
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
	 * exportExcel����-poi Excel����.
	 * ������������������SXSSFWorkbook�����ڴ����
	 * �������������ǳ�����������ʱ�ļ���wb.
	 * for 07Excel 
	 * ��׺ xlsx
	 * @param title
	 *            ����������
	 * @param dataset
	 *            ���������ݼ�
	 * @param out
	 *            �����
	 */
	@SuppressWarnings("unchecked")
	public void export07Excel(String title, Collection<T> dataset, OutputStream out) throws Exception {
		// ����һ��������
		// ���ȼ�����ݿ��Ƿ�����ȷ��
		Iterator<T> iterator = dataset.iterator();
		if (dataset == null || !iterator.hasNext() || title == null || out == null) {
			throw new Exception("��������ݲ������ԣ�");
		}
		// ȡ��ʵ�ʷ�����
		T tObject = iterator.next();
		Class<T> clazz = (Class<T>) tObject.getClass();

		//100Ϊ�������ڴ���������ÿ100��ˢϴһ�Σ������ڴ����
		SXSSFWorkbook workbook = new SXSSFWorkbook(100); 
		Sheet sheet = workbook.createSheet(title==null?"sheet1":title);
		// ���ñ��Ĭ���п��Ϊ20���ֽ�
		sheet.setDefaultColumnWidth(20);
		// ����һ����ʽ
		CellStyle style = workbook.createCellStyle();
		// ���ñ�����ʽ
		style = ExcelStyle.set07HeadStyle(workbook, style);
		// �õ������ֶ�
		Field filed[] = tObject.getClass().getDeclaredFields();

		// ����
		List<String> exportfieldtile = new ArrayList<String>();
		// �������ֶε�get����
		List<Method> methodObj = new ArrayList<Method>();
		
		// ��������filed
		for (int i = 0; i < filed.length; i++) {
			Field field = filed[i];
			ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
			// ���������annottion
			if (excelAnnotation != null) {
				String exprot = excelAnnotation.exportName();
				// ��ӵ�����
				exportfieldtile.add(exprot);
				// ��ӵ���Ҫ�������ֶεķ���
				String fieldname = field.getName();
				String getMethodName = "get" + fieldname.substring(0, 1).toUpperCase() + fieldname.substring(1);
				Method getMethod = clazz.getMethod(getMethodName, new Class[] {});
				methodObj.add(getMethod);
			}
		}

		// ������������
		Row row = sheet.createRow(0);
		for (int i = 0; i < exportfieldtile.size(); i++) {
			Cell cell = row.createCell(i);
			cell.setCellStyle(style);
			HSSFRichTextString text = new HSSFRichTextString(exportfieldtile.get(i));
			cell.setCellValue(text);
		}

		// ѭ����������
		int index = 0;
		iterator = dataset.iterator();
		while (iterator.hasNext()) {
			// �ӵڶ��п�ʼд����һ���Ǳ���
			index++;
			row = sheet.createRow(index);
			T t = (T) iterator.next();
			for (int k = 0; k < methodObj.size(); k++) {
				Cell cell = row.createCell(k);
				Method getMethod = methodObj.get(k);
				Object value = getMethod.invoke(t, new Object[] {});
				String textValue = getValue(value);
				cell.setCellValue(textValue);
			}
		}
		workbook.write(out);
		//��ʱ�ļ��Ĵ���
		workbook.dispose();
	}

	/**
	 * ���ؼ���poi����
	 * @param file ������ļ�
	 * @param pattern
	 * @return
	 */
	public Collection<T> importExcel(File file, String... pattern) throws Exception {
		
			//�ж���03����07��ʽ��excel
			String suffix = file.getPath().substring( file.getPath().lastIndexOf(".")+1);
			if(!"xls".equals(suffix) && !"xlsx".equals(suffix)){
				throw new Exception("�ļ���ʽ����excel���ļ���ʽ��"); 
			}
			Collection<T> dist = new ArrayList<T>();
			/**
			 * �෴��õ����÷���
			 */
			// �õ�Ŀ��Ŀ��������е��ֶ��б�
			Field[] fields = clazz.getDeclaredFields();
			// �����б���Annotation���ֶΣ�Ҳ�������������ݵ��ֶ�,���뵽һ��map��
			Map<String, Method> fieldMap = new HashMap<String, Method>();
			// ѭ����ȡ�����ֶ�
			for (Field field : fields) {
				// �õ������ֶ��ϵ�Annotation
				ExcelAnnotation excelAnnotation = field.getAnnotation(ExcelAnnotation.class);
				// �����ʶ��Annotationd
				if (excelAnnotation != null) {
					String fieldName = field.getName();
					// ����������Annotation���ֶε�Setter����
					String setMethodName = "set" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
					// ������õ�method
					Method setMethod = clazz.getMethod(setMethodName, new Class[] { field.getType() });
					// �����method��Annotaion������Ϊkey������
					fieldMap.put(excelAnnotation.exportName(), setMethod);
				}
			}
			
			Iterator<Row> row = null;

			/**
			 * excel�Ľ�����ʼ
			 */
			// �������File����ΪFileInputStream;
			FileInputStream inputStream = new FileInputStream(file);
			if("xlsx".equals(suffix)){
				//�ɶ�ȡ��������
				XSSFWorkbook wb = new XSSFWorkbook(inputStream);
				//��һ��sheet
				XSSFSheet sheet = wb.getSheetAt(0);
				row = sheet.rowIterator();
			}else{
				// �õ�������
				HSSFWorkbook book = new HSSFWorkbook(inputStream);
				// �õ���һҳ
				HSSFSheet sheet = book.getSheetAt(0);
				// �õ���һ���������
				row = sheet.rowIterator();
			}

			/**
			 * �������
			 */
			// �õ���һ�У�Ҳ���Ǳ�����
			Row titleRow = row.next();
			// �õ���һ�е�������
			Iterator<Cell> cellTitle = titleRow.cellIterator();
			// ��������������ݷ��뵽һ��map��
			Map<Integer, String> titleMap = new HashMap<Integer, String>();
			// �ӱ����һ�п�ʼ
			int i = 0;
			// ѭ���������е���
			while (cellTitle.hasNext()) {
				Cell cell = (Cell) cellTitle.next();
				String value = cell.getStringCellValue();
				titleMap.put(i, value);
				i++;
			}

			/**
			 * ����������
			 */
			while (row.hasNext()) {
				// �����µĵ�һ��
				Row rown = row.next();
				// �е�������
				Iterator<Cell> cellBody = rown.cellIterator();
				// �õ��������ʵ��
				T tObject = clazz.newInstance();
				// ����һ�е���
				int col = 0;
				while (cellBody.hasNext()) {
					Cell cell = (Cell) cellBody.next();
					// ����õ����еĶ�Ӧ�ı���
					String titleString = titleMap.get(col++);
					// �����һ�еı�������е�ĳһ�е�Annotation��ͬ����ô����ô���ĵ�set������������ֵ
					if (fieldMap.containsKey(titleString)) {
						Method setMethod = fieldMap.get(titleString);
						// �õ�setter�����Ĳ���
						Type[] types = setMethod.getGenericParameterTypes();
						// ֻҪһ������
						String xclass = String.valueOf(types[0]);
						// �жϲ�������
						if ("class java.lang.String".equals(xclass)) {
							setMethod.invoke(tObject, cell.getStringCellValue());
						} else if ("class java.util.Date".equals(xclass)) {
							// setMethod.invoke(tObject,
							// cell.getDateCellValue());
							setMethod.invoke(tObject, sdf.parse(cell.getStringCellValue()));
						} else if ("class java.lang.Boolean".equals(xclass) || "boolean".equals(xclass)) {	
							Boolean boolName = true;
							if ("��".equals(cell.getStringCellValue())) {
								boolName = false;
							}
							setMethod.invoke(tObject, boolName);
						} else if ("class java.lang.Integer".equals(xclass) || "int".equals(xclass)) {		//�������� �� ��װ����
							// setMethod.invoke(tObject, new
							// Integer(String.valueOf((int)cell.getNumericCellValue())));
							setMethod.invoke(tObject, new Integer(cell.getStringCellValue()));
						} else if ("class java.lang.Long".equals(xclass) || "long".equals(xclass)) {
							setMethod.invoke(tObject, new Long(cell.getStringCellValue()));
						} else if ("class java.lang.Double".equals(xclass) || "double".equals(xclass)) {
							setMethod.invoke(tObject, new Double(cell.getStringCellValue()));
						} else {
							//��Ҫ������ͣ����ﲹ��
						}
					}
				}
				dist.add(tObject);
			}
		return dist;
	}

	/**
	 * getValue����-cellֵ����.
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
			boolean bValue = (Boolean) value;
			textValue = "��";
			if (!bValue) {
				textValue = "��";
			}
		} else if (value instanceof Date) {
			Date date = (Date) value;
			textValue = sdf.format(date);

		} else {
			textValue = value.toString();
		}
		return textValue;
	}
	
}
