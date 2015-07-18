package cn.snt.util;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excelע�⣬����vo�б������뵼�������ݱ�ͷ
 * ��Ӣ�Ķ���
 * eg:@ExcelAnnotation(exportName="�Ա�")
 * @author hurong
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAnnotation {
	// excel����ʱ������ʾ�����֣����û������Annotation���ԣ������ᱻ�����͵���
	public String exportName();
}
