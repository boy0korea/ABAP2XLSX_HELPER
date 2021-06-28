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
  DATA: lr_data     TYPE REF TO data,
        lt_receiver TYPE TABLE OF string.
  FIELD-SYMBOLS: <lt_data> TYPE table.

  ASSIGN lr_data->* TO <lt_data>.


  CALL FUNCTION 'ZA2XH_EMAIL'
    EXPORTING
      it_data     = <lt_data>
*     it_field    = it_field
*     iv_filename = iv_filename
*     iv_subject  = iv_subject
*     iv_sender   = iv_sender
      it_receiver = lt_receiver.
ENDFORM.
