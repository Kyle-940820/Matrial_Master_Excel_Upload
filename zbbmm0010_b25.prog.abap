*&---------------------------------------------------------------------*
*& Report ZBBMM0010
*&---------------------------------------------------------------------*
*&   [MM]
*&   개발자        : CL2 kdt-b-25 하정훈
*&   프로그램 개요   : 자재 MASTER 업로드 프로그램
*&   개발 시작일    :'2024.10.28'
*&   개발 완료일    :'2024.10.30'
*&   개발상태      : 개발 완료
*&---------------------------------------------------------------------*
REPORT ZBBMM0010_B25 MESSAGE-ID ZCOMMON_MSG.

INCLUDE ZBBMM0010_B25_TOP.
INCLUDE ZBBMM0010_B25_S01.
INCLUDE ZBBMM0010_B25_O01.
INCLUDE ZBBMM0010_B25_I01.
INCLUDE ZBBMM0010_B25_F01.

INITIALIZATION.
  SSCRFIELDS-FUNCTXT_01 = ICON_XLS && '양식 다운로드'.

AT SELECTION-SCREEN ON VALUE-REQUEST FOR PA_PATH.
  PERFORM GET_FILE_PATH.

" '양식 다운로드' 버튼 눌렀을 때, EXCEL 파일 다운로드 후 실행.
AT SELECTION-SCREEN.
  CASE SSCRFIELDS-UCOMM.
    WHEN 'FC01'.
      PERFORM DOWN_FILE.
  ENDCASE.

START-OF-SELECTION.
  PERFORM CHECK_EXCEL. " 파일을 선택했는지 확인.
  PERFORM UPLOAD_EXCEL.
  CALL SCREEN 100.
