METHOD handledefault .
  IF zcl_abap2xlsx_helper=>is_abap2xlsx_installed( ) EQ abap_true.
    zcl_abap2xlsx_helper=>wd_upload_popup(
      EXPORTING
        iv_callback_action = 'DUMMY'
        io_view            = wd_this->wd_get_api( )
    ).
  ELSE.
    zcl_za2xh_upload_popup=>open_popup( NEW cl_fpm_parameter( ) ).
  ENDIF.
ENDMETHOD.

method WDDOEXIT .
endmethod.

method WDDOINIT .
endmethod.

method WDDOONCLOSE .
endmethod.

method WDDOONOPEN .
endmethod.

