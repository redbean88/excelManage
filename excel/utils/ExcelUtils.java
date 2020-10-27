package egovframework.com.excel.utils;

import java.util.List;

import org.springframework.web.multipart.commons.CommonsMultipartFile;

public interface ExcelUtils {

	public List excelReadSetValue(CommonsMultipartFile file, int strartRowNum, int startCelNum) throws Exception;
	
}
