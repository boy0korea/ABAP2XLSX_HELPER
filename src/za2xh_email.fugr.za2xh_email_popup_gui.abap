FUNCTION za2xh_email_popup_gui .
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     REFERENCE(IO_PARAM) TYPE REF TO  IF_FPM_PARAMETER
*"----------------------------------------------------------------------

  CREATE OBJECT go_assist.
  IF io_param IS NOT INITIAL.
    go_assist->mo_param = io_param.
  ELSE.
    CREATE OBJECT go_assist->mo_param TYPE cl_fpm_parameter.
  ENDIF.


  IF go_edit IS NOT INITIAL.
    go_edit->delete_text( ).
  ENDIF.

  CALL SCREEN 2100 STARTING AT 5 5.

ENDFUNCTION.
