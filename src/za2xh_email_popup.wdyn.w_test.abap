METHOD handledefault .
  DATA: lt_spfli TYPE TABLE OF spfli.

  SELECT *
    INTO TABLE lt_spfli
    FROM spfli.

  IF zcl_abap2xlsx_helper=>is_abap2xlsx_installed( ) EQ abap_true.
    zcl_abap2xlsx_helper=>excel_email( it_data = lt_spfli ).
  ELSE.
    zcl_za2xh_email_popup=>open_popup( NEW cl_fpm_parameter( ) ).
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

