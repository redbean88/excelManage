package egovframework.com.excel.download;

import java.io.BufferedReader;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.sql.Clob;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.view.AbstractView;

import egovframework.com.excel.utils.ExcelColumn;
import egovframework.rte.psl.dataaccess.util.EgovMap;

public class POIExcelView4 extends AbstractView {
	
	 /** The content type for an Excel response */
    private static final String CONTENT_TYPE_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

	/**
	 * 파일명 렌더러
	 */
	@Override
	protected void renderMergedOutputModel(Map<String, Object> model,
			HttpServletRequest request, HttpServletResponse response)
			throws Exception {
		
		 XSSFWorkbook workbook = new XSSFWorkbook();

	        setContentType(CONTENT_TYPE_XLSX);

	        buildExcelDocument(model, workbook, request, response);
	        
	        // Set the filename
	        String sFilename = "";
	        if(model.get("filename") != null){
	            sFilename = (String)model.get("filename");
	        }else if(request.getAttribute("filename") != null){
	            sFilename = (String)request.getAttribute("filename");
	        }else{
	            sFilename = "통계";
	         }

	        response.setContentType(getContentType());
	        
	        String header = request.getHeader("User-Agent");
	        sFilename = sFilename.replaceAll("\r","").replaceAll("\n","");
	        if(header.contains("MSIE") || header.contains("Trident") || header.contains("Chrome")){
	            sFilename = URLEncoder.encode(sFilename,"UTF-8").replaceAll("\\+","%20");
	            response.setHeader("Content-Disposition","attachment;filename="+sFilename+".xlsx;");
	        }else{
	            sFilename = new String(sFilename.getBytes("UTF-8"),"ISO-8859-1");
	            response.setHeader("Content-Disposition","attachment;filename=\""+sFilename + ".xlsx\"");
	        }
	        
	        // Flush byte array to servlet output stream.
	        ServletOutputStream out = response.getOutputStream();
	        out.flush();
	        workbook.write(out);
	        out.flush();
		
	}    
	/**
	 * TODO create mutil sheet option need
	 */
	@SuppressWarnings("unchecked")
	protected void buildExcelDocument(Map model, XSSFWorkbook wb, HttpServletRequest req, HttpServletResponse resp) throws Exception {
			List<String> sheetNm = (List<String>) model.get("sheetNm"); // 엑셀 시트 이름
			List<String> columnVarArr = getColumnVarArr(model.get("targetVO"));// 각 컬럼의 변수 이름
			List<String> columnArr =  (List<String>) model.get("header"); // 각 컬럼 이름
			List<EgovMap> dataList = (List<EgovMap>) model.get("body"); // 데이터가 담긴 리스트 
			
			CellStyle cellStyle = headerStyler(wb);
			CellStyle cellStyle2 = bodyStyler(wb);
			
			for (int i = 0; i < sheetNm.size(); i++) {
				XSSFSheet sheet = wb.createSheet(sheetNm.get(i));
				sheet.setDefaultColumnWidth(12);
				
				processHeader(columnArr, dataList, cellStyle, cellStyle2, sheet);
				processBody(columnVarArr, columnArr, dataList, cellStyle2, sheet);
			}
			
    }
	


	/**
	 * Convenient method to set header.
	 * @param columnArr
	 * @param dataList
	 * @param cellStyle
	 * @param cellStyle2
	 * @param sheet
	 */
	private void processHeader(List<String> columnArr, List<?> dataList, CellStyle cellStyle, CellStyle cellStyle2, XSSFSheet sheet) {
		XSSFCell cell;
		for(int i=0; i<columnArr.size(); i++){
            setText(getCell(sheet, 0, i), columnArr.get(i));
            getCell(sheet, 0, i).setCellStyle(cellStyle);
            sheet.autoSizeColumn(i);
            int columnWidth = (sheet.getColumnWidth(i))*5;
            sheet.setColumnWidth(i, columnWidth);
            
            if(dataList.size() < 1){
                cell = getCell(sheet, 1, i);
                if(i==0){
                    setText(cell, "등록된 정보가 없습니다.");
                }
                cell.setCellStyle(cellStyle2);
            }
        }
	}
	
	/**
     * Convenient method to set body.
	 * @param columnVarArr
	 * @param columnArr
	 * @param dataList
	 * @param cellStyle2
	 * @param sheet
	 * @throws Exception 
	 */
	private void processBody(List<String> columnVarArr, List<String> columnArr,	List<EgovMap> dataList, CellStyle cellStyle2, XSSFSheet sheet) throws Exception {
		XSSFCell cell;
		if(dataList.size() > 0){ // 저장된 데이터가 있을때
            // 리스트 데이터 삽입
            for (int i = 0; i<dataList.size(); i++) {
                EgovMap dataEgovMap = dataList.get(i);
                
                // 맨 앞 컬럼인 "번호"는 idx라는 이름으로 여기서 생성하여 넣어준다.
                dataEgovMap.put("idx", (i+1)+""); 
                
                for(int j=0; j<columnVarArr.size(); j++){
                	if(j >= columnArr.size() ) break;	//제목이 없는값은 미표시
                    String data = dataParsing2String(dataEgovMap.get(columnVarArr.get(j)));
                    cell = getCell(sheet, 1 + i, j);
                    setText(cell, data);
                    cell.setCellStyle(cellStyle2);
                }
            }
        }else{ // 저장된 데이터가 없으면 셀 병합
            // 셀 병합(시작열, 종료열, 시작행, 종료행)
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, columnArr.size()-1));
        }
	}
	
	/**
	 * string 외의 타입을 string 타입으로 파싱한다
	 * @param object
	 * @return
	 */
	private String dataParsing2String(Object object) throws Exception {
		if(object instanceof java.sql.Clob){
			return clob2String((java.sql.Clob)object);
		}else{
			return object != null ? object.toString() : "";
		}
	}

	/**
	 * clob > String
	 * @param object
	 * @return
	 * @throws Exception 
	 */
	private String clob2String(Clob clob) throws Exception {
		StringBuffer strOut = new StringBuffer();
		String str = "";
		BufferedReader br = new BufferedReader(clob.getCharacterStream());
		while ((str = br.readLine()) != null) {
			strOut.append(str);
		}
		return strOut.toString();
	}
	/**
	 * Convenient method to get ColumnVarArr using reflection.
	 * <p>get field attached annotation ( eg.@ExcelColumn) </p>
	 * 
	 * @param target
	 * @return
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 */
	private List<String> getColumnVarArr(Object target) throws IllegalArgumentException, IllegalAccessException {
		 Map<Integer,String> resultMap = new HashMap<Integer,String>();
		 List<String> resultArray = new ArrayList<String>();
		for (Field field: target.getClass().getDeclaredFields()) {
			field.setAccessible(true);
			if ( field.isAnnotationPresent(ExcelColumn.class)){
				int key = field.getAnnotation(ExcelColumn.class).order();
				String value = (String) field.getName();
				if(key != -1) resultMap.put(key, value);
			}
		}
		
		for (int i = 0; i < resultMap.size(); i++) {
			resultArray.add(resultMap.get(i));
		}
		return resultArray;
	}

	/**
	 * Convenient method to set a cell style of body.
	 * @param wb
	 * @return
	 */
	private CellStyle bodyStyler(XSSFWorkbook wb) {
		CellStyle cellStyle2 = wb.createCellStyle(); // 데이터셀의 셀스타일
        cellStyle2.setWrapText(true); // 줄 바꿈           
        cellStyle2.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 셀 세로 정렬
        cellStyle2.setDataFormat((short)0x31); // 셀 데이터 형식
        cellStyle2.setBorderRight(XSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderLeft(XSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderTop(XSSFCellStyle.BORDER_THIN);
        cellStyle2.setBorderBottom(XSSFCellStyle.BORDER_THIN);
		return cellStyle2;
	}
    /**
     * Convenient method to set a cell style of header.
     * 
     * @param wb
     */
	private CellStyle headerStyler(XSSFWorkbook wb) {
		CellStyle cellStyle = wb.createCellStyle(); // 제목셀의 셀스타일
        cellStyle.setWrapText(true); // 줄 바꿈            
        cellStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index); // 셀 색상
        cellStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND); // 셀 색상 패턴
        cellStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER); // 셀 가로 정렬
        cellStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER); // 셀 세로 정렬
        cellStyle.setDataFormat((short)0x31); // 셀 데이터 형식
        cellStyle.setBorderRight(XSSFCellStyle.BORDER_DOUBLE);
        cellStyle.setBorderLeft(XSSFCellStyle.BORDER_DOUBLE);
        cellStyle.setBorderTop(XSSFCellStyle.BORDER_DOUBLE);
        cellStyle.setBorderBottom(XSSFCellStyle.BORDER_DOUBLE);
        
        // 셀 폰트색상, bold처리
        Font font = wb.createFont();
        font.setColor(HSSFColor.WHITE.index);
        font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
        cellStyle.setFont(font);
		return cellStyle;
	}

	 /**
     * Convenient method to obtain the cell in the given sheet, row and column.
     * 
     * <p>Creates the row and the cell if they still doesn't already exist.
     * Thus, the column can be passed as an int, the method making the needed downcasts.</p>
     * 
     * @param sheet a sheet object. The first sheet is usually obtained by workbook.getSheetAt(0)
     * @param row thr row number
     * @param col the column number
     * @return the XSSFCell
     */
    protected XSSFCell getCell(XSSFSheet sheet, int row, int col) {
        XSSFRow sheetRow = sheet.getRow(row);
        if (sheetRow == null) {
            sheetRow = sheet.createRow(row);
        }
        XSSFCell cell = sheetRow.getCell((short) col);
        if (cell == null) {
            cell = sheetRow.createCell((short) col);
        }
        return cell;
    }

    /**
     * Convenient method to set a String as text content in a cell.
     * 
     * @param cell the cell in which the text must be put
     * @param text the text to put in the cell
     */
    protected void setText(XSSFCell cell, String text) {
        cell.setCellType(XSSFCell.CELL_TYPE_STRING);
        cell.setCellValue(text);
    }



}
