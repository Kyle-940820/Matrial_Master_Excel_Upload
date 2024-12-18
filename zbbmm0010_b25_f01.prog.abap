*&---------------------------------------------------------------------*
*& Include          ZBBMM0010_F01
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Form GET_FILE_PATH
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM GET_FILE_PATH .
  " 파일 오픈 METHOD 호출
  CALL METHOD CL_GUI_FRONTEND_SERVICES=>FILE_OPEN_DIALOG
  " 확장자 EXCEL 로만 선택 가능.
    EXPORTING
      DEFAULT_EXTENSION = CL_GUI_FRONTEND_SERVICES=>FILETYPE_EXCEL
      FILE_FILTER       = CL_GUI_FRONTEND_SERVICES=>FILETYPE_EXCEL
    CHANGING
      FILE_TABLE        = GT_FILE
      RC                = GV_RC_FILEPATH.

  " 파일을 정상적으로 선택했으면, WA에 할당.
  IF GT_FILE IS NOT INITIAL.
    READ TABLE GT_FILE INTO GS_FILE INDEX 1.
    PA_PATH = GS_FILE-FILENAME.
  ELSE.
    MESSAGE I022 DISPLAY LIKE 'E'.
  ENDIF.

ENDFORM.
*&---------------------------------------------------------------------*
*& Form DOWN_FILE
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM DOWN_FILE .
  " EXCEL FILE 다운로드 로직에 필요한 변수 선언.
  DATA : LV_FILENAME    TYPE STRING,
         LV_PATH        TYPE STRING,
         LV_FULLPATH    TYPE STRING,
         LV_USER_ACTION TYPE I,
         LV_RC          TYPE SY-SUBRC,
         LS_KEY         TYPE WWWDATATAB.

  " SMW0 에 등록한 EXCEL에 대한 WWWDATA DATA를 LS_KEY에 할당.
  SELECT SINGLE * FROM WWWDATA
    INTO CORRESPONDING FIELDS OF LS_KEY
   WHERE OBJID = 'ZBMM1010_XLS_TEMPLATE'.

  " 파일 이름 자동 생성.
  CONCATENATE '자재 MASTER_' SY-DATUM SY-UZEIT INTO LV_FILENAME.

  " 확장자명, 파일명 받기.
  CALL METHOD CL_GUI_FRONTEND_SERVICES=>FILE_SAVE_DIALOG
    EXPORTING
      DEFAULT_EXTENSION         = 'XLSX'
      DEFAULT_FILE_NAME         = LV_FILENAME
    CHANGING
      FILENAME                  = LV_FILENAME
      PATH                      = LV_PATH
      FULLPATH                  = LV_FULLPATH
      USER_ACTION               = LV_USER_ACTION
    EXCEPTIONS
      CNTL_ERROR                = 1
      ERROR_NO_GUI              = 2
      NOT_SUPPORTED_BY_GUI      = 3
      INVALID_DEFAULT_FILE_NAME = 4
      OTHERS                    = 5.

  " 사용자가 다운로드 중 취소했을 때 EXIT.
  IF LV_USER_ACTION = CL_GUI_FRONTEND_SERVICES=>ACTION_CANCEL.
    EXIT.
  ENDIF.

  CALL FUNCTION 'DOWNLOAD_WEB_OBJECT'
    EXPORTING
      KEY         = LS_KEY
      DESTINATION = CONV LOCALFILE( LV_FULLPATH ).

  " LV_FULLPATH 에 할당한 SMW0 EXCEL 실행.
  CALL METHOD CL_GUI_FRONTEND_SERVICES=>EXECUTE
    EXPORTING
      DOCUMENT               = LV_FULLPATH
    EXCEPTIONS
      CNTL_ERROR             = 1
      ERROR_NO_GUI           = 2
      BAD_PARAMETER          = 3
      FILE_NOT_FOUND         = 4
      PATH_NOT_FOUND         = 5
      FILE_EXTENSION_UNKNOWN = 6
      ERROR_EXECUTE_FAILED   = 7
      SYNCHRONOUS_FAILED     = 8
      NOT_SUPPORTED_BY_GUI   = 9
      OTHERS                 = 10.

  IF SY-SUBRC = 0.
    MESSAGE S027.
  ELSE.
    MESSAGE I026 DISPLAY LIKE 'E'.
  ENDIF.

ENDFORM.
*&---------------------------------------------------------------------*
*& Form CHECK_EXCEL
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM CHECK_EXCEL .
  " PA_PATH PARAMETER에 파일을 선택하지 않았을 때 알람.
  DATA : LV_FILE LIKE RLGRAP-FILENAME.

  LV_FILE = PA_PATH+3.

  IF LV_FILE = SPACE.
    MESSAGE I022 DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form UPLOAD_EXCEL
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM UPLOAD_EXCEL .

  " EXCEL DATA를 INTERNAL TABLE에 담기.
  CALL FUNCTION 'ALSM_EXCEL_TO_INTERNAL_TABLE'
    EXPORTING
      FILENAME                = PA_PATH
      I_BEGIN_COL             = 2
      I_BEGIN_ROW             = 3
      I_END_COL               = 100
      I_END_ROW               = 10000
    TABLES
      INTERN                  = GT_EXCEL
    EXCEPTIONS
      INCONSISTENT_PARAMETERS = 1
      UPLOAD_OLE              = 2
      OTHERS                  = 3.

  " EXCEL DATA 가져오기에 실패했을 때.
  IF SY-SUBRC <> 0.
    MESSAGE I029 DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
  ENDIF.

  " EXCEL DATA 에 DATA가 없을 때.
  IF GT_EXCEL IS INITIAL.
    MESSAGE I030 DISPLAY LIKE 'E'.
    LEAVE LIST-PROCESSING.
    ROLLBACK WORK.
  ENDIF.

  " EXCEL DATA를 한 줄씩 GS_DATA에 담고 GT_DATA에 생성.
  LOOP AT GT_EXCEL INTO GS_EXCEL.
    ASSIGN COMPONENT GS_EXCEL-COL OF STRUCTURE GS_DATA TO <FS_COMP>.
    <FS_COMP> = GS_EXCEL-VALUE.

    AT END OF ROW.
      APPEND GS_DATA TO GT_DATA.
      CLEAR: GS_DATA.
    ENDAT.
  ENDLOOP.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form CREATE_OBJECT
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM CREATE_OBJECT .
  IF GO_DOCK IS INITIAL.
    CREATE OBJECT GO_DOCK
      EXPORTING
        REPID                       = SY-REPID
        DYNNR                       = SY-DYNNR
        SIDE                        = CL_GUI_DOCKING_CONTAINER=>DOCK_AT_TOP
        EXTENSION                   = 1000
      EXCEPTIONS
        CNTL_ERROR                  = 1
        CNTL_SYSTEM_ERROR           = 2
        CREATE_ERROR                = 3
        LIFETIME_ERROR              = 4
        LIFETIME_DYNPRO_DYNPRO_LINK = 5
        OTHERS                      = 6.
    IF SY-SUBRC <> 0.
*     MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*                WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.

    CREATE OBJECT GO_ALV
      EXPORTING
        I_PARENT          = GO_DOCK
      EXCEPTIONS
        ERROR_CNTL_CREATE = 1
        ERROR_CNTL_INIT   = 2
        ERROR_CNTL_LINK   = 3
        ERROR_DP_CREATE   = 4
        OTHERS            = 5.
    IF SY-SUBRC <> 0.
*     MESSAGE ID SY-MSGID TYPE SY-MSGTY NUMBER SY-MSGNO
*                WITH SY-MSGV1 SY-MSGV2 SY-MSGV3 SY-MSGV4.
    ENDIF.

    PERFORM SET_LAYOUT.
    PERFORM SET_SORT.
    PERFORM SET_FCAT.

    GS_VARIANT-REPORT = SY-CPROG.
    GS_VARIANT-VARIANT = '/LAYOUT1'.

    CALL METHOD GO_ALV->SET_TABLE_FOR_FIRST_DISPLAY
      EXPORTING
*       i_buffer_active               =
        I_BYPASSING_BUFFER            = 'X'
*       i_consistency_check           =
        I_STRUCTURE_NAME              = 'ZSBMM1010_STR'
        IS_VARIANT                    = GS_VARIANT
        I_SAVE                        = 'A'
*       i_default                     = 'X'
        IS_LAYOUT                     = GS_LAYO
*       is_print                      =
*       it_special_groups             =
*       it_toolbar_excluding          =
*       it_hyperlink                  =
*       it_alv_graphics               =
*       it_except_qinfo               =
*       ir_salv_adapter               =
      CHANGING
        IT_OUTTAB                     = GT_DATA[]
        IT_FIELDCATALOG               = GT_FCAT
        IT_SORT                       = GT_SORT
*       it_filter                     =
      EXCEPTIONS
        INVALID_PARAMETER_COMBINATION = 1
        PROGRAM_ERROR                 = 2
        TOO_MANY_LINES                = 3
        OTHERS                        = 4.
    IF SY-SUBRC <> 0.
*     Implement suitable error handling here
    ENDIF.
  ENDIF.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form SET_LAYOUT
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM SET_LAYOUT .
  GS_LAYO-CWIDTH_OPT = 'A'.
  GS_LAYO-GRID_TITLE = 'Excel 자재 데이터 리스트'.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form SET_SORT
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM SET_SORT .
*  " '자재유형' 필드를 기준으로 SORT.
*  GS_SORT-SPOS = 1.
*  GS_SORT-FIELDNAME = 'MATTYPE'.
*  GS_SORT-UP = 'X'.
*  APPEND GS_SORT TO GT_SORT.
*  CLEAR GS_SORT.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form SET_FCAT
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM SET_FCAT .

*           MATNAME   TYPE ZTBMM1011-MATNAME,
*         MATTYPE   TYPE  ZTBMM1010-MATTYPE,
*         PRDTYPE   TYPE ZTBMM1010-PRODTYPE,
*         MSQUAN    TYPE ZTBMM1010-MSQUAN,
*         UNITCODE1 TYPE ZTBMM1010-UNITCODE1,
*         WEIGHT    TYPE  ZTBMM1010-WEIGHT,
*         UNITCODE2 TYPE ZTBMM1010-UNITCODE2,
*         VOLUME    TYPE  ZTBMM1010-VOLUME,
*         UNITCODE3 TYPE ZTBMM1010-UNITCODE3,


  GS_FCAT-FIELDNAME = 'MATNAME'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'MATTYPE'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'PRDTYPE'.
  GS_FCAT-JUST = 'C'.
  GS_FCAT-COLTEXT = '완제품 유형'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'MSQUAN'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'UNITCODE1'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'WEIGHT'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'UNITCODE2'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'VOLUME'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'UNITCODE3'.
  GS_FCAT-JUST = 'C'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.


  GS_FCAT-FIELDNAME = 'MATCODE'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'SPRAS'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'ALL'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'VAL'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'DEL'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_DATE_F'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_TIME_F'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_USER_F'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_DATE_L'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_TIME_L'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'STAMP_USER_L'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.

  GS_FCAT-FIELDNAME = 'DELFLG'.
  GS_FCAT-NO_OUT = 'X'.
  APPEND GS_FCAT TO GT_FCAT.
  CLEAR GS_FCAT.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form CONFIRM_POPUP
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*&      <-- LV_CHK
*&---------------------------------------------------------------------*
FORM CONFIRM_POPUP.
  DATA : LV_ANSWER.

  CALL FUNCTION 'POPUP_TO_CONFIRM'
    EXPORTING
      TEXT_QUESTION         = TEXT-Q01
      TEXT_BUTTON_1         = 'YES'
      ICON_BUTTON_1         = 'ICON_OKAY'
      TEXT_BUTTON_2         = 'NO'
      ICON_BUTTON_2         = 'ICON_CANCEL'
      DEFAULT_BUTTON        = '1'
      DISPLAY_CANCEL_BUTTON = ''
    IMPORTING
      ANSWER                = LV_ANSWER.

  IF LV_ANSWER = '2'.
    LEAVE TO SCREEN 0.
  ENDIF.
  CHECK LV_ANSWER = '1'.
ENDFORM.
*&---------------------------------------------------------------------*
*& Form SAVE_EXCEL
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM SAVE_EXCEL .
  DATA: LV_NUM     TYPE NRIV,    " MATCODE 채번.
        LV_MATCODE TYPE STRING,  " MATCODE 채번 시, 'MAT' CONCAT.
        LV_COUNT   TYPE I.       " SELECT 문에서 카운트 변수.

  "EXCEL 데이터가 현재 GT_DATA INTERNAL TABLE에 있으므로, GT_ZTBMM1010과 GT_ZTBMM1011에 옮겨줘야 함.
  MOVE-CORRESPONDING GT_DATA TO GT_ZTBMM1010.
  MOVE-CORRESPONDING GT_DATA TO GT_ZTBMM1011.

  " GT_ZTBMM1010 TABLE에 MATCODE, 타임스탬프 값 입력.
  LOOP AT GT_ZTBMM1010 INTO GS_ZTBMM1010.

    " MATCODE 채번.
    CALL FUNCTION 'NUMBER_GET_NEXT'
      EXPORTING
        NR_RANGE_NR             = '01'
        OBJECT                  = 'ZBBMM1010'
      IMPORTING
        NUMBER                  = LV_NUM
      EXCEPTIONS
        INTERVAL_NOT_FOUND      = 1
        NUMBER_RANGE_NOT_INTERN = 2
        OBJECT_NOT_FOUND        = 3
        QUANTITY_IS_0           = 4
        QUANTITY_IS_NOT_1       = 5
        INTERVAL_OVERFLOW       = 6
        BUFFER_OVERFLOW         = 7
        OTHERS                  = 8.
    IF SY-SUBRC = 0.
      CONCATENATE 'MAT' LV_NUM INTO LV_MATCODE.
    ENDIF.

    " 채번된 MATCODE WORK AREA에 할당.
    GS_ZTBMM1010-MATCODE = LV_MATCODE.

    " 타임스탬프 WORK AREA에 할당.
    SELECT SINGLE EMPID
     INTO GS_ZTBMM1010-STAMP_USER_F
     FROM ZTBSD1030
     WHERE LOGID = SY-UNAME.

    GS_ZTBMM1010-STAMP_DATE_F = SY-DATUM.
    GS_ZTBMM1010-STAMP_TIME_F = SY-UZEIT.

    " 마스터 테이블 TYPE INTERNAL TABLE에 WORK AREA 데이터 할당.
    MODIFY GT_ZTBMM1010 FROM GS_ZTBMM1010 INDEX SY-TABIX.
  ENDLOOP.

  " 마스터 테이블 UPDATE.
  IF SY-SUBRC = 0.
    MODIFY ZTBMM1010 FROM TABLE GT_ZTBMM1010.
  ENDIF.

  " GT_ZTBMM1011 TABLE에 MATCODE, SPRAS, 타임스탬프 값 입력.
  LOOP AT GT_ZTBMM1011 INTO GS_ZTBMM1011.
    READ TABLE GT_ZTBMM1010 INTO GS_ZTBMM1010 INDEX SY-TABIX.

    GS_ZTBMM1011-MATCODE = GS_ZTBMM1010-MATCODE.
    GS_ZTBMM1011-SPRAS = SY-LANGU.

    SELECT SINGLE EMPID
     INTO GS_ZTBMM1011-STAMP_USER_F
     FROM ZTBSD1030
     WHERE LOGID = SY-UNAME.

    GS_ZTBMM1011-STAMP_DATE_F = SY-DATUM.
    GS_ZTBMM1011-STAMP_TIME_F = SY-UZEIT.

    MODIFY GT_ZTBMM1011 FROM GS_ZTBMM1011 INDEX SY-TABIX.
  ENDLOOP.

  " TEXT 테이블 UPDATE.
  IF SY-SUBRC = 0.
    MODIFY ZTBMM1011 FROM TABLE GT_ZTBMM1011.

    "MATCODE 값 가져오기.
    MOVE-CORRESPONDING GT_ZTBMM1010 TO GT_ZTBMM1011.
    IF SY-SUBRC = 0.
      MESSAGE I031 DISPLAY LIKE 'S'.
    ELSE.
      MESSAGE I032 DISPLAY LIKE 'E'.
    ENDIF.
  ENDIF.

  CLEAR: LV_MATCODE, LV_NUM, GS_ZTBMM1010, GS_ZTBMM1011.
ENDFORM.
