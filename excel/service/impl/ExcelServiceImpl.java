package egovframework.com.excel.service.impl;

import java.util.List;
import java.util.Map;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;

import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;
import org.springframework.ui.ModelMap;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.multipart.commons.CommonsMultipartFile;
import org.springframework.web.servlet.ModelAndView;

import egovframework.com.excel.download.POIExcelView4;
import egovframework.com.excel.service.ExcelService;
import egovframework.com.excel.utils.ExcelUtils;
import egovframework.com.excel.utils.ExcelxlsView;
import egovframework.rte.fdl.cmmn.EgovAbstractServiceImpl;
import egovframework.rte.fdl.property.EgovPropertyService;

@Service("excelService")
public class ExcelServiceImpl extends EgovAbstractServiceImpl implements ExcelService{
	
    /** 공통프로퍼티 */
	@Resource(name = "propertiesService")
	protected EgovPropertyService propertiesService;
	
	/** 엑셀 유틸 */
	@Resource(name = "poiExcelUtil")
	protected ExcelUtils poiExcelUtil;
    
    Logger logger = LoggerFactory.getLogger(ExcelServiceImpl.class);
    
	/**
	 * 엑셀출력
	 * @param originFileNm
	 * @param model
	 * @param list
	 * @return
	 * @throws Exception
	 */
	public ModelAndView templateDownload(String originFileNm, ModelMap model, String fileNm) throws Exception {
		return processPrint(originFileNm, model, fileNm);
	}
	public ModelAndView templateDownload(ModelMap model, String fileNm) throws Exception {
		String originFileNm = "defualTemplate"; 
		return processPrint(originFileNm, model, fileNm);
	}


	/**
	 * @param originFileNm
	 * @param model
	 * @param fileNm TODO
	 * @return
	 */
	private ModelAndView processPrint(String originFileNm, ModelMap model, String fileNm) {
		if(StringUtils.isBlank(fileNm)) fileNm = "통계현황";
		String extend = ".xls";
		String fileStorePath = propertiesService.getString("Globals.excelFileStorePath");
		
		model.addAttribute("templateFilePath", fileStorePath + originFileNm + extend);
		model.addAttribute("downFileName", fileNm + extend);
		
		return new ModelAndView(new ExcelxlsView(), model);
	}


	/**
	 * 엑셀업로드
	 * @param request
	 * @return
	 * @throws Exception
	 */
	@SuppressWarnings("unchecked")
	public List<List<List<?>>> getExcelData(HttpServletRequest request) throws Exception {
		
		MultipartHttpServletRequest multiRequest = (MultipartHttpServletRequest)request;
		CommonsMultipartFile file = (CommonsMultipartFile)multiRequest.getFile("excelFile");		//파일 정보
		
		//엑셀정보
		int strartRowNum = 1;	//2번째 행(row)부터 읽음
		int startCelNum = 0; 	//1번째 열(cell)부터 읽음
		return poiExcelUtil.excelReadSetValue(file, strartRowNum, startCelNum);
	}
	/**
	 * 단순 엑셀 다운로드
	 */
	public ModelAndView simpleDownload(Map data)
			throws Exception {
		return new ModelAndView(new POIExcelView4(), data);
	}
}
