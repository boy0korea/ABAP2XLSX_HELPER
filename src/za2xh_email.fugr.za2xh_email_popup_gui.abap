FUNCTION za2xh_email_popup_gui .
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     REFERENCE(IO_EVENT_DATA) TYPE REF TO  IF_FPM_PARAMETER
*"----------------------------------------------------------------------
  DATA: lt_receiver TYPE TABLE OF string.

  IF go_assist IS INITIAL.
    CREATE OBJECT go_assist.
  ENDIF.

  IF io_event_data IS NOT INITIAL.
    go_assist->mo_event_data = io_event_data.
    io_event_data->get_value(
      EXPORTING
        iv_key   = 'IT_RECEIVER'
      IMPORTING
        ev_value = lt_receiver
    ).
    CONCATENATE LINES OF lt_receiver INTO gv_email SEPARATED BY cl_abap_char_utilities=>newline.
  ELSE.
    CREATE OBJECT go_assist->mo_event_data TYPE cl_fpm_parameter.
  ENDIF.

  IF go_edit IS NOT INITIAL.
    go_edit->delete_text( ).
  ENDIF.

  CALL SCREEN 2100 STARTING AT 5 5.

ENDFUNCTION.
