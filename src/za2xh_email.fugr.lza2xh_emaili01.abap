*----------------------------------------------------------------------*
***INCLUDE LZA2XH_EMAILI01.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*&      Module  USER_COMMAND_2100  INPUT
*&---------------------------------------------------------------------*
*       text
*----------------------------------------------------------------------*
MODULE user_command_2100 INPUT.
  CASE sy-ucomm.
    WHEN 'OK'.
      PERFORM on_ok.
    WHEN OTHERS.
      LEAVE TO SCREEN 0.
  ENDCASE.
ENDMODULE.