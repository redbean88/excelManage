package egovframework.com.excel.service;

import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;

import org.springframework.stereotype.Service;
import org.springframework.ui.ModelMap;
import org.springframework.web.servlet.ModelAndView;

@Service
public interface ExcelService{
    
	/**
	 * 엑셀 츌력(템플릿 이용)
	 * @param originFileNm
	 * @param model
	 * @param fileNm TODO
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public ModelAndView templateDownload(String originFileNm, ModelMap model, String fileNm) throws Exception;
	public ModelAndView templateDownload(ModelMap model, String fileNm) throws Exception;

	
	/**
	 * 엑셀 다운로드
	 * @param data
	 * @return
	 * @throws Exception
	 */
	public ModelAndView simpleDownload(Map data ) throws Exception;
	
	/**
	 * 엑셀 데이터 획득
	 * @param request
	 * @return
	 * @throws Exception
	 */
	public List<List<List<?>>> getExcelData(HttpServletRequest request) throws Exception;

}
