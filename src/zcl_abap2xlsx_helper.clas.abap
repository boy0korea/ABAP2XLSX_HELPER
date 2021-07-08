class ZCL_ABAP2XLSX_HELPER definition
  public
  create public .

public section.

  types:
    BEGIN OF ts_field,
        fieldname    TYPE fieldname,
        label_text   TYPE scrtext_l,
        fixed_values TYPE wdr_context_attr_value_list,
      END OF ts_field .
  types:
    tt_field TYPE TABLE OF ts_field .

  class-methods EXCEL_DOWNLOAD
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods EXCEL_EMAIL
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_SUBJECT type CLIKE optional
      !IV_SENDER type CLIKE optional
      !IT_RECEIVER type STRINGTAB optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L' .
  class-methods EXCEL_UPLOAD
    importing
      !IV_EXCEL type XSTRING optional
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_BEGIN_ROW type INT4 default 2
      !IV_SHEET_NO type INT1 default 1
    exporting
      !ET_DATA type STANDARD TABLE
      !EV_ERROR_TEXT type STRING .
  class-methods GET_FIELDCATALOG
    importing
      !IT_DATA type STANDARD TABLE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !ET_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD .
  class-methods CONVERT_ABAP_TO_EXCEL
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods CONVERT_JSON_TO_EXCEL
    importing
      !IV_DATA_JSON type STRING
      !IT_DDIC_OBJECT type DD_X031L_TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods CONVERT_EXCEL_TO_ABAP
    importing
      !IV_EXCEL type XSTRING
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_BEGIN_ROW type INT4 default 2
      !IV_SHEET_NO type INT1 default 1
    exporting
      !ET_DATA type STANDARD TABLE
      !EV_ERROR_TEXT type STRING .
  class-methods GET_XSTRING_FROM_SMW0
    importing
      !IV_SMW0 type WWWDATA-OBJID
    returning
      value(RV_XSTRING) type XSTRING .
  class-methods FPM_UPLOAD_POPUP
    importing
      !IV_CALLBACK_EVENT_ID type FPM_EVENT_ID default 'ZA2XH_UPLOAD' .
  class-methods WD_UPLOAD_POPUP
    importing
      !IV_CALLBACK_ACTION type STRING
      !IO_VIEW type ref to IF_WD_VIEW_CONTROLLER .
  class-methods DEFAULT_EXCEL_FILENAME
    returning
      value(RV_FILENAME) type STRING .
  class-methods GET_DDIC_FIXED_VALUES
    importing
      !IO_TYPE type ref to CL_ABAP_TYPEDESCR
    returning
      value(RT_DDL) type WDR_CONTEXT_ATTR_VALUE_LIST .
  class-methods TEST .
  class-methods IS_ABAP2XLSX_INSTALLED
    importing
      !IV_WITH_MESSAGE type FLAG default ABAP_TRUE
    returning
      value(RV_INSTALLED) type FLAG .
  class-methods MESSAGE
    importing
      !IV_ERROR_TEXT type CLIKE .
  class-methods CHECK_INSTALL
    importing
      !IV_CLASS_NAME type CLIKE
      !IV_ERROR_TEXT type CLIKE optional
    returning
      value(RV_INSTALLED) type FLAG .
  PROTECTED SECTION.

    CLASS-METHODS readme .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ABAP2XLSX_HELPER IMPLEMENTATION.


  METHOD check_install.
* https://github.com/boy0korea/ABAP_INSTALL_CHECK

    TRY.
        cl_abap_typedescr=>describe_by_name(
          EXPORTING
            p_name         = iv_class_name
          EXCEPTIONS
            type_not_found = 1
        ).
      CATCH cx_root.
        " error
        sy-subrc = 4.
    ENDTRY.

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
        it_data                 = it_data
        it_field                = it_field
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr
      IMPORTING
        ev_excel                = ev_excel
        ev_error_text           = ev_error_text.
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


  METHOD convert_json_to_excel.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('CONVERT_JSON_TO_EXCEL')
*    CALL METHOD zcl_abap2xlsx_helper_int=>convert_json_to_excel
      EXPORTING
        iv_data_json            = iv_data_json
        it_ddic_object          = it_ddic_object
        it_field                = it_field
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr
      IMPORTING
        ev_excel                = ev_excel
        ev_error_text           = ev_error_text.
  ENDMETHOD.


  METHOD default_excel_filename.
    rv_filename = |{ sy-uname }_{ sy-datum }_{ sy-uzeit }.xlsx|.
  ENDMETHOD.


  METHOD excel_download.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('EXCEL_DOWNLOAD')
*    CALL METHOD zcl_abap2xlsx_helper_int=>excel_download
      EXPORTING
        it_data                 = it_data
        it_field                = it_field
        iv_filename             = iv_filename
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr
      IMPORTING
        ev_excel                = ev_excel
        ev_error_text           = ev_error_text.
  ENDMETHOD.


  METHOD excel_email.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('EXCEL_EMAIL')
*    CALL METHOD zcl_abap2xlsx_helper_int=>excel_email
      EXPORTING
        it_data                 = it_data
        it_field                = it_field
        iv_subject              = iv_subject
        iv_sender               = iv_sender
        it_receiver             = it_receiver
        iv_filename             = iv_filename
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr.
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


  METHOD fpm_upload_popup.
    DATA: lo_event_data TYPE REF TO if_fpm_parameter.

    CREATE OBJECT lo_event_data TYPE cl_fpm_parameter.

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_CALLBACK_EVENT_ID'
        iv_value = iv_callback_event_id
    ).

    zcl_za2xh_upload_popup=>open_popup( lo_event_data ).
  ENDMETHOD.


  METHOD get_ddic_fixed_values.
    DATA: lt_fixed_value TYPE ddfixvalues,
          ls_fixed_value TYPE ddfixvalue,
          ls_ddl         TYPE wdr_context_attr_value.

    IF io_type IS INSTANCE OF cl_abap_elemdescr.
      CAST cl_abap_elemdescr( io_type )->get_ddic_fixed_values(
*          EXPORTING
*            p_langu        = SY-LANGU       " Current Language
        RECEIVING
          p_fixed_values = lt_fixed_value " Defaults
        EXCEPTIONS
          not_found      = 1              " Type could not be found
          no_ddic_type   = 2              " Typ is not a dictionary type
          OTHERS         = 3
      ).

      LOOP AT lt_fixed_value INTO ls_fixed_value WHERE option = 'EQ'.
        CLEAR: ls_ddl.
        ls_ddl-value = ls_fixed_value-low.
        ls_ddl-text = ls_fixed_value-ddtext.
        APPEND ls_ddl TO rt_ddl.
      ENDLOOP.
    ENDIF.

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


  METHOD get_xstring_from_smw0.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('GET_XSTRING_FROM_SMW0')
*    CALL METHOD zcl_abap2xlsx_helper_int=>GET_XSTRING_FROM_SMW0
      EXPORTING
        iv_smw0    = iv_smw0
      RECEIVING
        rv_xstring = rv_xstring.
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
        iv_image_xstring = zcl_abap2xlsx_helper=>get_xstring_from_smw0( 'S_F_FAVO' )
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


    zcl_abap2xlsx_helper=>excel_email(
      EXPORTING
        it_data                 = lt_sflight
        it_field                = lt_field
    ).

  ENDMETHOD.


  METHOD wd_upload_popup.
    DATA: lo_event_data TYPE REF TO if_fpm_parameter.

    CREATE OBJECT lo_event_data TYPE cl_fpm_parameter.

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_CALLBACK_ACTION'
        iv_value = iv_callback_action
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IO_VIEW'
        iv_value = CAST cl_wdr_view( io_view )
    ).

    zcl_za2xh_upload_popup=>open_popup( lo_event_data ).
  ENDMETHOD.
ENDCLASS.
