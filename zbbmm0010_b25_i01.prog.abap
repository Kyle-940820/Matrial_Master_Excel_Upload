*&---------------------------------------------------------------------*
*& Include          ZBBMM0010_I01
*&---------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  EXIT  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE EXIT INPUT.
  CASE OK_CODE.
    WHEN 'EXIT'.
      LEAVE PROGRAM.
    WHEN 'CANCEL'.
      LEAVE TO SCREEN 0.
  ENDCASE.
ENDMODULE.
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_0100  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE USER_COMMAND_0100 INPUT.
  DATA: LV_CHK.

  CASE OK_CODE.
    WHEN 'BACK'.
      LEAVE TO SCREEN 0.
    WHEN 'SAVE'.
      " CONFIRM POPUP 에서 '예'를 눌렀는지 확인 하고 SAVE_EXCEL SUBROUTINE 실행.
      PERFORM CONFIRM_POPUP.
      PERFORM SAVE_EXCEL.
      LEAVE TO SCREEN 0.
  ENDCASE.
ENDMODULE.
