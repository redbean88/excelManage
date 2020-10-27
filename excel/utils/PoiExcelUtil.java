package egovframework.com.excel.utils;

import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.commons.CommonsMultipartFile;

@Service("poiExcelUtil")
public class PoiExcelUtil implements ExcelUtils {

	/**
	 * 엑셀 파일을 읽는다
	 */
	public List<List<List<Object>>> excelReadSetValue(CommonsMultipartFile file, int strartRowNum, int startCelNum) throws Exception {
		//xls, xlsx 구분
		Workbook workbook = processExtend(file);
		
		if (workbook.getNumberOfSheets() > 0) return processSheet(workbook,strartRowNum,startCelNum);
		return null;
	}
	/**
	 * 시트별 처리
	 * @param workbook
	 * @param strartRowNum
	 * @param startCelNum
	 * @return
	 */
	private List<List<List<Object>>> processSheet(Workbook workbook, int strartRowNum, int startCelNum) {
		List<List<List<Object>>> resultList = new ArrayList<>();
		
		for (int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
			resultList.add(processRows(workbook, sheetIdx, strartRowNum, startCelNum));
		}
		return resultList;
	}
	/**
	 * 각 행 처리
	 * @param workbook
	 * @param sheetIdx
	 * @param strartRowNum
	 * @param startCelNum
	 * @return
	 */
	private List<List<Object>>processRows(Workbook workbook, int sheetIdx ,int strartRowNum, int startCelNum) {
		List<List<Object>> resultList = new ArrayList<>();
		//Sheet 선택
		Sheet sheet = workbook.getSheetAt(sheetIdx);
		
		for(int row = strartRowNum ; row < sheet.getPhysicalNumberOfRows(); row++) {
			resultList.add(processCells(sheet, row, strartRowNum, startCelNum));
		}
		return resultList;
	}
	/**
	 * 각 열 처리
	 * @param sheet
	 * @param idx
	 * @param strartRowNum
	 * @param startCelNum
	 * @return
	 */
	private List<Object> processCells(Sheet sheet, int idx, int strartRowNum,int startCelNum) {
			List<Object> resultList = new ArrayList<>(); 
		
			//한 줄씩 읽고 데이터 저장
			Row row = sheet.getRow(idx);
			int maxCell = 49;
			int lastCellIdx = sheet.getRow(0).getPhysicalNumberOfCells() > maxCell ? sheet.getRow(0).getPhysicalNumberOfCells() : maxCell;
			
			//Cell 기본값 빼고 시작(0에서 시작)
			for(int cellIdx = startCelNum ; cellIdx < lastCellIdx; cellIdx++) {
				if(cellIdx < sheet.getRow(0).getPhysicalNumberOfCells()){
					resultList.add(processValues(row.getCell(cellIdx)));
				}else{
					resultList.add("");
				}
			}
		return resultList;
	}

	/**
	 * 각 값처리
	 * @param cell
	 * @return
	 */
	private String processValues( Cell cell) {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
		String value = "";
		switch(cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK :
			value = "";
			break;
		case Cell.CELL_TYPE_BOOLEAN :
			value = "" + cell.getBooleanCellValue();
			break;
		case Cell.CELL_TYPE_ERROR :
			value = "" + cell.getErrorCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA :
			value = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_NUMERIC :
			if(HSSFDateUtil.isInternalDateFormat(cell.getCellStyle().getDataFormat())) {
				value = sdf.format(cell.getDateCellValue());
			}
			else {
				cell.setCellType(Cell.CELL_TYPE_STRING ); 
				value = cell.getStringCellValue(); 
			}
			break;
		case Cell.CELL_TYPE_STRING :
			value = cell.getStringCellValue();
			break;
		}
		//공백과 트림 제거
		value = value.trim().replaceAll(" ", "");
		return value;
	}

	/**
	 * @param file
	 * @return
	 * @throws IOException
	 */
	private Workbook processExtend(CommonsMultipartFile file)
			throws IOException {
		Workbook workbook;
		if(file.getOriginalFilename().toUpperCase().endsWith("XLSX")) {
			workbook = new XSSFWorkbook(file.getInputStream());
		}
		else {
			workbook = new HSSFWorkbook(file.getInputStream());
		}
		return workbook;
	}
	
    protected void buildExcelDocument(Map model, XSSFWorkbook wb, HttpServletRequest req, HttpServletResponse resp) throws Exception {
        XSSFCell cell = null;
 
        String sheetNm = (String) model.get("sheetNm"); // 엑셀 시트 이름
        
        String[] columnArr = (String[]) model.get("columnArr"); // 각 컬럼 이름
        List<?> dataList = (List<?>) model.get("body"); // 데이터가 담긴 리스트 
        
        CellStyle cellStyle = headerStyler(wb);
        CellStyle cellStyle2 = bodyStyler(wb);
        
        XSSFSheet sheet = wb.createSheet(sheetNm);
        sheet.setDefaultColumnWidth(12);
        
        processHeader(columnArr, dataList, cellStyle, cellStyle2, sheet);
        processBody(columnArr, dataList, cellStyle2, sheet);
    }
	/**
	 * @param columnArr
	 * @param dataList
	 * @param cellStyle
	 * @param cellStyle2
	 * @param sheet
	 */
	private void processHeader(String[] columnArr, List<?> dataList,
			CellStyle cellStyle, CellStyle cellStyle2, XSSFSheet sheet) {
		XSSFCell cell;
		for(int i=0; i<columnArr.length; i++){
            setText(getCell(sheet, 0, i), columnArr[i]);
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
	 * @param columnArr
	 * @param dataList
	 * @param cellStyle2
	 * @param sheet
	 * @throws IllegalAccessException
	 */
	private void processBody(String[] columnArr, List<?> dataList,
			CellStyle cellStyle2, XSSFSheet sheet)
			throws IllegalAccessException {
		XSSFCell cell;
		if(dataList.size() > 0){ // 저장된 데이터가 있을때
            // 리스트 데이터 삽입
            for (int i = 0; i<dataList.size(); i++) {
                Object target = dataList.get(i);
                
            	for (Field field: target.getClass().getDeclaredFields()) {
        			field.setAccessible(true);
        			if ( field.isAnnotationPresent(ExcelColumn.class)){
        				int idx = field.getAnnotation(ExcelColumn.class).order();
        				cell = getCell(sheet, 1 + i, idx);
        				setText(cell, (String) field.get(target));
        				cell.setCellStyle(cellStyle2);
        			}
        		}
            	
            }
        }else{ // 저장된 데이터가 없으면 셀 병합
            // 셀 병합(시작열, 종료열, 시작행, 종료행)
            sheet.addMergedRegion(new CellRangeAddress(1, 1, 0, columnArr.length-1));
        }
	}
	/**
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
	 * @param wb
	 * @return
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
