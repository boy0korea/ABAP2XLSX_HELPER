*----------------------------------------------------------------------*
***INCLUDE LZA2XH_EMAILO01.
*----------------------------------------------------------------------*
*&---------------------------------------------------------------------*
*& Module STATUS_2100 OUTPUT
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
MODULE status_2100 OUTPUT.
  SET PF-STATUS '2100'.
  SET TITLEBAR '2100'.

  IF go_edit IS INITIAL.
    CREATE OBJECT go_cont
      EXPORTING
        container_name = 'TEXT_EDIT'.
    CREATE OBJECT go_edit
      EXPORTING
        parent = go_cont.
  ENDIF.

  IF gv_email IS NOT INITIAL.
    go_edit->set_textstream( gv_email ).
    CLEAR: gv_email.
  ENDIF.

  go_edit->set_focus( go_edit ).

ENDMODULE.
