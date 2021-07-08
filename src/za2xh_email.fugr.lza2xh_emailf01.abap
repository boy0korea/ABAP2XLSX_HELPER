*----------------------------------------------------------------------*
***INCLUDE LZA2XH_EMAILF01.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Form on_ok
*&---------------------------------------------------------------------*
*& text
*&---------------------------------------------------------------------*
*& -->  p1        text
*& <--  p2        text
*&---------------------------------------------------------------------*
FORM on_ok .
  DATA: lv_text TYPE string.

  go_edit->get_textstream(
*    EXPORTING
*      only_when_modified     = false       " get text only when modified
    IMPORTING
      text                   = lv_text        " Text as String with Carriage Returns and Linefeeds
*      is_modified            = is_modified " modify status of text
*    EXCEPTIONS
*      error_cntl_call_method = 1           " Error while retrieving a property from TextEdit control
*      not_supported_by_gui   = 2           " Method is not supported by installed GUI
*      others                 = 3
  ).
  cl_gui_cfw=>flush( ).

  go_assist->mo_param->set_value(
    EXPORTING
      iv_key   = 'IT_RECEIVER'
      iv_value = go_assist->split_email_string( lv_text )
  ).

  go_assist->on_ok( ).
  LEAVE TO SCREEN 0.
ENDFORM.
