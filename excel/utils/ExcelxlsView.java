package egovframework.com.excel.utils;

import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import net.sf.jxls.transformer.XLSTransformer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.servlet.view.AbstractView;


/**
 * 엑셀 수정 다운로드
 * 수정하기 워힌 원본 엑섹 파일이 필요
 * @author User
 *
 */
public class ExcelxlsView extends AbstractView {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelxlsView.class);
	 
	@SuppressWarnings({ "unchecked", "rawtypes" })
	@Override
	protected void renderMergedOutputModel( 
			Map model, 
			HttpServletRequest request, 
			HttpServletResponse response 
			) throws Exception {
		// 엑셀 텔플릿 가져오기
		String templateFilePath = (String) model.get("templateFilePath");
		// 저장될 파일명
		String downFileName = (String) model.get("downFileName");
        
        //파일 정보 제거
        model.remove("xlsTemplateFilePath");
        model.remove("xlsDownFileName");
		
        XLSTransformer transformer = new XLSTransformer();
        
        BufferedInputStream fin = null;
    	
    	try{
	        fin = new BufferedInputStream(new FileInputStream(templateFilePath));
	        
	        HSSFWorkbook workbook = (HSSFWorkbook) transformer.transformXLS(fin, model);
			
	        response.setHeader("Content-Disposition", "attachment; filename=\"" + URLEncoder.encode(downFileName, "UTF-8") + "\"");
			response.setHeader("Content-Transfer-Encoding", "binary;");
			response.setHeader("Pragma", "no-cache");
			response.setHeader("Expires", "-1");
	        workbook.write(response.getOutputStream());
			response.getOutputStream().flush();
			response.getOutputStream().close();
    	}catch(Exception e){
    		if(templateFilePath == null) throw new Exception("템플렛 파일을 찾을수 없습니다.");
    		else e.printStackTrace();
    	}finally{
    		if (fin != null) {
    			try {
    			    fin.close();
    			} catch (Exception ignore) {
    			    //System.out.println("IGNORED: " + ignore.getMessage());
    				LOGGER.debug("IGNORED: " + ignore.getMessage());
    			}
    		}
    	}
	}	
}
