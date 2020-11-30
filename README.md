# excelManage

# dependency

```xml
<!-- 엑셀 제어(xls) maven setting -->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi</artifactId>
			<version>3.9</version>
		</dependency>
<!-- 엑셀 제어(xlsx) maven setting-->
		<dependency>
			<groupId>org.apache.poi</groupId>
			<artifactId>poi-ooxml</artifactId>
			<version>3.9</version>
		</dependency>
```

# 사용법
+ 대상 객체 어노테이션 추가
```
public class ContractMngVO {

	@ExcelColumn(order=2)
	private String cntiType;					//계약 구분(수주 E, 자체 발주 O, 수주 발주 EO)

	@ExcelColumn(order=1)
	private String cntiTitle;						//계약명
	
	@ExcelColumn(order=3)
	private String cntiPrice;					//계약 금액

	@ExcelColumn(order=5)
	private String cntiOrdrStDt;				//계약 시작일
	@ExcelColumn(order=6)
	private String cntiOrdrEdDt;				//계약 종료일

	@ExcelColumn(order=8)
	private String cntiEtcMemo;				//기타사항
	
	private String cntiCorpId;					//계약업체고유키
	@ExcelColumn(order=0)
	private String cntiCorpNm;				//계약업체명

	@ExcelColumn(order=7)
	private String cntiCorpUserNm2;			//계약 담당자명2
	@ExcelColumn(order=4)
	private String cntiDate;						//계약일
	

  ```
  |어노테이션|기능|
  |:--:|:--:|
  |ExcelColumn(order=[num])| 엑셀 출력 대상 어노테이션 ( order : 출력 행 번호)|
  
  + 컨트롤러
  ```
  @RequestMapping(value="/projectmng/contract/excelList.do")
	public ModelAndView exceldown(ModelMap model ,
			@ModelAttribute("contractMngVO") ContractMngVO contractMngVO
			) throws Exception{
		contractMngVO.setRecordCountPerPage(999999999);

		
		List<String> sheetList = new ArrayList<String>();
		sheetList.add("통계");
		List<String> headerList = new ArrayList<String>();
		headerList.add("고객사");
		headerList.add("계약명");
		headerList.add("계약종류");
		headerList.add("계약금액");
		headerList.add("계약일");
		headerList.add("계약기간");
		headerList.add("계약기간");
		headerList.add("실무담당자");
		headerList.add("기타사항");
		
		model.addAttribute("fileNm","통계");
		model.addAttribute("sheetNm",sheetList);
		model.addAttribute("header",headerList);
		model.addAttribute("targetVO",new ContractMngVO());
		model.addAttribute("body",processContractType(contractMngService.listContract(contractMngVO)));
		
		return new ModelAndView(new POIExcelView4(), model);
	}
  ```
    |어노테이션|기능|
  |:--:|:--:|
  |sheetList| 시트명|
  |headerList| 엑셀 컬럼 명 리스트|
  |targetVO| 엑셀 출력 대상 객체|
  |body| 엑셀 출력 대상 객체 리스트|
  | new ModelAndView(new POIExcelView4(), model);| 실체 엑셀 출력 |
  
  
  
  
