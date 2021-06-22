CLASS zcl_abap2xlsx_helper DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    TYPES:
      BEGIN OF ts_field,
        fieldname    TYPE fieldname,
        label_text   TYPE scrtext_l,
        fixed_values TYPE wdr_context_attr_value_list,
      END OF ts_field .
    TYPES:
      tt_field TYPE TABLE OF ts_field .

    CLASS-METHODS excel_download
      IMPORTING
        !it_data              TYPE STANDARD TABLE
        !it_field             TYPE zcl_abap2xlsx_helper=>tt_field OPTIONAL
        !iv_filename          TYPE clike OPTIONAL
        !iv_sheet_title       TYPE clike OPTIONAL
        !iv_auto_column_width TYPE flag DEFAULT abap_true
        !iv_default_descr     TYPE c DEFAULT 'L'
      EXPORTING
        !ev_excel             TYPE xstring
        !ev_error_text        TYPE string .
    CLASS-METHODS excel_upload
      IMPORTING
        !iv_excel      TYPE xstring OPTIONAL
        !it_field      TYPE zcl_abap2xlsx_helper=>tt_field OPTIONAL
        !iv_begin_row  TYPE int4 DEFAULT 2
        !iv_sheet_no   TYPE int1 DEFAULT 1
      EXPORTING
        !et_data       TYPE STANDARD TABLE
        !ev_error_text TYPE string .
    CLASS-METHODS get_fieldcatalog
      IMPORTING
        !it_data          TYPE STANDARD TABLE
        !iv_default_descr TYPE c DEFAULT 'L'
      EXPORTING
        !et_field         TYPE zcl_abap2xlsx_helper=>tt_field .
    CLASS-METHODS convert_abap_to_excel
      IMPORTING
        !it_data              TYPE STANDARD TABLE
        !it_field             TYPE zcl_abap2xlsx_helper=>tt_field OPTIONAL
        !iv_sheet_title       TYPE clike OPTIONAL
        !iv_auto_column_width TYPE flag DEFAULT abap_true
        !iv_default_descr     TYPE c DEFAULT 'L'
      EXPORTING
        !ev_excel             TYPE xstring
        !ev_error_text        TYPE string .
    CLASS-METHODS convert_excel_to_abap
      IMPORTING
        !iv_excel      TYPE xstring
        !it_field      TYPE zcl_abap2xlsx_helper=>tt_field OPTIONAL
        !iv_begin_row  TYPE int4 DEFAULT 2
        !iv_sheet_no   TYPE int1 DEFAULT 1
      EXPORTING
        !et_data       TYPE STANDARD TABLE
        !ev_error_text TYPE string .
    CLASS-METHODS test .
    CLASS-METHODS is_abap2xlsx_installed
      IMPORTING
        !iv_with_message    TYPE flag DEFAULT abap_true
      RETURNING
        VALUE(rv_installed) TYPE flag .
    CLASS-METHODS message
      IMPORTING
        !iv_error_text TYPE clike .
    CLASS-METHODS check_install
      IMPORTING
        !iv_class_name      TYPE clike
        !iv_error_text      TYPE clike OPTIONAL
      RETURNING
        VALUE(rv_installed) TYPE flag .
  PROTECTED SECTION.

    CLASS-METHODS readme .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ABAP2XLSX_HELPER IMPLEMENTATION.


  METHOD check_install.
* https://github.com/boy0korea/ABAP_INSTALL_CHECK

    cl_abap_typedescr=>describe_by_name(
      EXPORTING
        p_name         = iv_class_name
      EXCEPTIONS
        type_not_found = 1
    ).

    IF sy-subrc EQ 0.
      " exist
      rv_installed = abap_true.
    ELSEIF iv_error_text IS NOT INITIAL.
      " not exist
      message( iv_error_text = iv_error_text ).
    ENDIF.
  ENDMETHOD.


  METHOD convert_abap_to_excel.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('CONVERT_ABAP_TO_EXCEL')
*    CALL METHOD zcl_abap2xlsx_helper_int=>convert_abap_to_excel
      EXPORTING
        it_data              = it_data
        it_field             = it_field
        iv_sheet_title       = iv_sheet_title
        iv_auto_column_width = iv_auto_column_width
        iv_default_descr     = iv_default_descr
      IMPORTING
        ev_excel             = ev_excel
        ev_error_text        = ev_error_text.
  ENDMETHOD.


  METHOD convert_excel_to_abap.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('CONVERT_EXCEL_TO_ABAP')
*    CALL METHOD zcl_abap2xlsx_helper_int=>convert_excel_to_abap
      EXPORTING
        iv_excel      = iv_excel
        it_field      = it_field
        iv_begin_row  = iv_begin_row
        iv_sheet_no   = iv_sheet_no
      IMPORTING
        et_data       = et_data
        ev_error_text = ev_error_text.
  ENDMETHOD.


  METHOD excel_download.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('EXCEL_DOWNLOAD')
*    CALL METHOD zcl_abap2xlsx_helper_int=>excel_download
      EXPORTING
        it_data              = it_data
        it_field             = it_field
        iv_filename          = iv_filename
        iv_sheet_title       = iv_sheet_title
        iv_auto_column_width = iv_auto_column_width
        iv_default_descr     = iv_default_descr
      IMPORTING
        ev_excel             = ev_excel
        ev_error_text        = ev_error_text.
  ENDMETHOD.


  METHOD excel_upload.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('EXCEL_UPLOAD')
*    CALL METHOD zcl_abap2xlsx_helper_int=>excel_upload
      EXPORTING
        iv_excel      = iv_excel
        it_field      = it_field
        iv_begin_row  = iv_begin_row
        iv_sheet_no   = iv_sheet_no
      IMPORTING
        et_data       = et_data
        ev_error_text = ev_error_text.
  ENDMETHOD.


  METHOD get_fieldcatalog.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('GET_FIELDCATALOG')
*    CALL METHOD zcl_abap2xlsx_helper_int=>get_fieldcatalog
      EXPORTING
        it_data          = it_data
        iv_default_descr = iv_default_descr
      IMPORTING
        et_field         = et_field.
  ENDMETHOD.


  METHOD is_abap2xlsx_installed.
    DATA: lv_class_name TYPE string VALUE 'ZCL_EXCEL'.

    check_install(
      EXPORTING
        iv_class_name =  lv_class_name
      RECEIVING
        rv_installed  = rv_installed
    ).
    IF rv_installed EQ abap_false AND iv_with_message EQ abap_true.
      AUTHORITY-CHECK OBJECT 'S_DEVELOP' ID 'ACTVT' FIELD '03'.
      IF sy-subrc EQ 0.
        " for developer
        message( 'install abap2xlsx from https://github.com/sapmentors/abap2xlsx' ).
      ELSE.
        " for user
        message( 'abap2xlsx is not installed.' ).
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD message.
    CHECK: iv_error_text IS NOT INITIAL.

    IF wdr_task=>application IS NOT INITIAL.
      " WD or FPM
      wdr_task=>application->component->if_wd_controller~get_message_manager( )->report_error_message(
        EXPORTING
          message_text = iv_error_text
      ).
    ELSE.
      " GUI
      MESSAGE iv_error_text TYPE 'S' DISPLAY LIKE 'E'.
    ENDIF.

  ENDMETHOD.


  METHOD readme.
* https://github.com/boy0korea/ABAP2XLSX_HELPER
  ENDMETHOD.


  METHOD test.
    DATA: lt_sflight  TYPE TABLE OF sflight,
          lt_sflight2 TYPE TABLE OF sflight,
          lt_field    TYPE zcl_abap2xlsx_helper=>tt_field,
          ls_sflight  TYPE sflight,
          lv_xstring  TYPE xstring.
    FIELD-SYMBOLS: <ls_field> TYPE zcl_abap2xlsx_helper=>ts_field.

    SELECT *
      FROM sflight
      INTO TABLE lt_sflight.

    IF lt_sflight IS INITIAL.
*      CALL TRANSACTION 'BC_DATA_GEN' AND SKIP FIRST SCREEN.
      SUBMIT sapbc_data_generator USING SELECTION-SET 'SAP&BC_MINI' WITH pa_dark = abap_true AND RETURN.
      SELECT *
        FROM sflight
        INTO TABLE lt_sflight.
    ENDIF.

    zcl_abap2xlsx_helper=>get_fieldcatalog(
      EXPORTING
        it_data          = lt_sflight
      IMPORTING
        et_field         = lt_field
    ).
    LOOP AT lt_field ASSIGNING <ls_field>.
      CASE <ls_field>-fieldname.
        WHEN 'CARRID'.
          SELECT carrid AS value carrname AS text
            INTO CORRESPONDING FIELDS OF TABLE <ls_field>-fixed_values
            FROM scarr.
          SORT <ls_field>-fixed_values BY value.
        WHEN 'CONNID'.
          SELECT connid AS value connid AS text
            INTO CORRESPONDING FIELDS OF TABLE <ls_field>-fixed_values
            FROM spfli.
          SORT <ls_field>-fixed_values BY value.
        WHEN 'CURRENCY'.
          SELECT currkey AS value currkey AS text
            INTO CORRESPONDING FIELDS OF TABLE <ls_field>-fixed_values
            FROM scurx.
          SORT <ls_field>-fixed_values BY value.
        WHEN 'PLANETYPE'.
          SELECT planetype AS value planetype AS text
            INTO CORRESPONDING FIELDS OF TABLE <ls_field>-fixed_values
            FROM saplane.
          SORT <ls_field>-fixed_values BY value.
        WHEN OTHERS.
      ENDCASE.
    ENDLOOP.

    zcl_abap2xlsx_helper=>excel_download(
      EXPORTING
        it_data  = lt_sflight
        it_field = lt_field
      IMPORTING
        ev_excel = lv_xstring
    ).

    zcl_abap2xlsx_helper=>excel_upload(
      EXPORTING
        iv_excel     = lv_xstring
      IMPORTING
        et_data      = lt_sflight2
    ).
    ls_sflight-mandt = sy-mandt.
    MODIFY lt_sflight2 FROM ls_sflight TRANSPORTING mandt WHERE mandt <> ls_sflight-mandt.

    IF lt_sflight <> lt_sflight2.
      BREAK-POINT.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
