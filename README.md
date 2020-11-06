# excelManage

# 사용방법

### 적용절차
+ 상세페이지조회 컨트롤러
  + 어노테이션 및 파라미터 추가
  + @NoticeInsertAnnotation
  + @ModelAttribute("noticeVO") NoticeVO noticeVO 
  + **taskId**는 업무 PK값, **taskManageUserId**는 사용자 PK값
  > 예제
  ```java
  
    @NoticeInsertAnnotation
	@RequestMapping("/maintenance/selectTaskMngView.do")
	public String SelectTaskMngView(@ModelAttribute("maintenanceVO") MaintenanceVO maintenanceVO, 
			@ModelAttribute("noticeVO") NoticeVO noticeVO 	,//new마크 표시를 위한 처리
			ModelMap model) throws Exception {
      
      ...업무프로세스...
      
      
      };
   ```
   
+ new 표시 갱신 컨트롤러
  + 어노테이션 및 파라미터 추가
  + @@NoticeUpdateTimeOfContentAnnotation
  + @ModelAttribute("noticeVO") NoticeVO noticeVO 	
  + **taskId**는 업무 PK값, **taskManageUserId**는 사용자 PK값
  > 예제
  ```java
  //업무수정처리
	@NoticeUpdateTimeOfContentAnnotation //new마크 표시를 위한 처리
	@RequestMapping("/maintenance/updateTaskMngEditProc.do")
	public String UpdateTaskMngEditProc(@ModelAttribute("maintenanceVO") MaintenanceVO maintenanceVO, 
			@ModelAttribute("noticeVO") NoticeVO noticeVO 	,//new마크 표시를 위한 처리
			ModelMap model, EmplyrVo emplyrVO) throws Exception {
      
        ...업무프로세스...
      
      };
  
  ```
+ 목록 조회 쿼리
  + new마커 표시 쿼리 추가
  ```sql
  CASE WHEN 
  NVL((select 
    LAST_UPDUSR_PNTTM 
  from 
    NOTICE_MNG_COUNTER C 
  where 
    C.CONTER_FLAG = 'contents' AND C.MNG_UNIQ_ID = E.MNG_UNIQ_ID
  ),E.MNG_EDIT_DATE) 																		/*해당글 최종 수정일(댓글 포함)*/
   >
          NVL(
          (select 
            /*+ INDEX_DESC(C MAINTENACE_CNT_IDX)*/ LAST_UPDUSR_PNTTM 
          from
            NOTICE_MNG_COUNTER C 
          where 
            C.CONTER_FLAG = 'user' AND C.MNG_UNIQ_ID = E.MNG_UNIQ_ID AND C.UNIQ_ID = #taskManageUserId# AND ROWNUM = 1
  ),TO_DATE ('20000101','YYYYMMDD'))  																		/*해당글 최종 확인일*/
          THEN 'new'
             ELSE 'old' END
          AS NEW_MARK 
  ```
