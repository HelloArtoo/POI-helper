package cn.snt.util;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * excel注解，用于vo中标明导入导出的数据表头
 * 中英文对照
 * eg:@ExcelAnnotation(exportName="性别")
 * @author hurong
 *
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelAnnotation {
	// excel导出时标题显示的名字，如果没有设置Annotation属性，将不会被导出和导入
	public String exportName();
}
