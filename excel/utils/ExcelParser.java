package egovframework.com.excel.utils;

import java.lang.reflect.Field;
import java.util.List;

public class ExcelParser{
	private static ExcelParser excelParser = new ExcelParser();
	
	private ExcelParser(){}
	
   public static ExcelParser getInstance(){
        return excelParser;

    }

	public static Object List2Target(List<?> uploadData , Object target) throws Exception {
		
		for (Field field: target.getClass().getDeclaredFields()) {
			field.setAccessible(true);
			if ( field.isAnnotationPresent(ExcelColumn.class)){
				int setIdx = field.getAnnotation(ExcelColumn.class).order();
				if(setIdx > (uploadData.size()-1)) throw new RuntimeException("ExcelColumn 값이 리스트의 길이보다 큽니다.");
				if(setIdx != -1){
					field.set(target, uploadData.get(setIdx));
				}
			}
		}
		return target;
	}
}
