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

  class-data GV_LIST_UIBB_EXPORT_ON type FLAG value ABAP_TRUE ##NO_TEXT.

  class-methods EXCEL_DOWNLOAD
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
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
  class-methods ENH_CL_FPM_LIST_UIBB_ASSIST_AT
    importing
      !IV_FORMAT type FPMGB_EXPORT_FORMAT
      !IRT_RESULT_DATA type ref to DATA
      !IT_P13N_COLUMN type IF_FPM_LIST_SETTINGS_VARIANT=>TY_T_O_COLUMN
      !IT_FIELD_USAGE type FPMGB_T_FIELDUSAGE .
  class-methods ENH_CL_FPM_LIST_UIBB_RENDERER_
    importing
      !IO_EXPORT_BTN_CHOICE type ref to CL_WD_TOOLBAR_BTN_CHOICE
      !IO_VIEW type ref to IF_WD_VIEW .
protected section.

  class-methods README .
private section.
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


  METHOD enh_cl_fpm_list_uibb_assist_at.
    DATA: lt_field2      TYPE tt_field,
          lt_field       TYPE tt_field,
          ls_field       TYPE ts_field,
          ls_field_usage TYPE fpmgb_s_fieldusage,
          lo_p13n_column TYPE REF TO if_fpm_list_settings_column,
          lv_column_name TYPE string.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    CHECK: is_abap2xlsx_installed( ) EQ abap_true,
           gv_list_uibb_export_on EQ abap_true.

    IF iv_format EQ 'ZA2X'.
      ASSIGN irt_result_data->* TO <lt_data>.

      get_fieldcatalog(
        EXPORTING
          it_data          = <lt_data>
        IMPORTING
          et_field         = lt_field2
      ).

      LOOP AT it_p13n_column INTO lo_p13n_column.
        CHECK: lo_p13n_column->is_visible( ).
        lv_column_name = lo_p13n_column->get_name( ).
        READ TABLE lt_field2 INTO ls_field WITH KEY fieldname = lv_column_name.
        CHECK: sy-subrc EQ 0.
        READ TABLE it_field_usage INTO ls_field_usage WITH KEY name = lv_column_name.
        ls_field-label_text = ls_field_usage-label_text.
        APPEND ls_field TO lt_field.
      ENDLOOP.

      excel_download(
        EXPORTING
          it_data              = <lt_data>
          it_field             = lt_field
      ).
    ENDIF.


* enhancement 위치:
*Enhanced Development Object    CL_FPM_LIST_UIBB_ASSIST_ATS
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""$"$\SE:(1) Class LCL_EXPORT_ACTION, Method EXECUTE, Start                                                                                                    A
*$*$-Start: (1)---------------------------------------------------------------------------------$*$*
*ENHANCEMENT 1  ZE_ABAP2XLSX_HELPER_LIST_ASSIS.    "active version
** additional export menu.
*    DATA: zlrt_result_data TYPE REF TO data.
*
*    IF me->mv_format EQ 'ZA2X'.
*      me->get_result_data(
*        exporting
*          iv_data_only        = abap_true
*        importing
*          ert_result_data     = zlrt_result_data
*      ).
*      zcl_abap2xlsx_helper=>enh_cl_fpm_list_uibb_assist_at(
*        EXPORTING
*          iv_format       = me->mv_format
*          irt_result_data = zlrt_result_data
*          it_p13n_column  = me->mo_list_uibb_assist->mo_personalization_api->get_current_variant( )->get_columns( )
*          it_field_usage  = me->mo_list_uibb_assist->mt_field_usage
*      ).
*      RETURN.
*    ENDIF.
*ENDENHANCEMENT.
*$*$-End:   (1)---------------------------------------------------------------------------------$*$*
  ENDMETHOD.


  METHOD enh_cl_fpm_list_uibb_renderer_.
    DATA: lo_tab_action TYPE REF TO cl_wd_menu_action_item.

    CHECK: is_abap2xlsx_installed( ) EQ abap_true,
           gv_list_uibb_export_on EQ abap_true.

    CHECK: io_export_btn_choice IS BOUND.

*     io_export_btn_choice->remove_choice(
*       EXPORTING
*         id         = 'MNUAI_FPM_EXPORT_CSV'
*     ).
*     io_export_btn_choice->remove_choice(
*       EXPORTING
*         id         = 'MNUAI_FPM_EXPORT_PDF'
*     ).

    lo_tab_action =
      cl_wd_menu_action_item=>new_menu_action_item(
        id           = `ZMNUAI_FPM_EXPORT_ZDB`              "#EC NOTEXT
        view         = io_view
        on_action    = 'DISPATCH_EXPORT'  " lif_renderer_constants=>cs_table_action-export
        text         = 'Excel in DB format'
        enabled      = abap_true
        visible      = abap_true
    ).
    DATA(lt_action_parameters) = VALUE wdr_name_value_list(
      (
         name  = 'FORMAT'   " lc_export_action_format_param
         value = 'ZA2X'
      )
    ).
    lo_tab_action->map_on_action( lt_action_parameters ).
    io_export_btn_choice->add_choice( lo_tab_action ).

* enhancement 위치:
*Enhanced Development Object    CL_FPM_LIST_UIBB_RENDERER_ATS
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""$"$\SE:(1) Class LCL_TABLE_RENDERER, Method RENDER_STANDARD_TOOLBAR_ITEMS, End                                                                               A
*$*$-Start: (1)---------------------------------------------------------------------------------$*$*
*ENHANCEMENT 1  ZE_ABAP2XLSX_HELPER_LIST_RENDE.    "active version
** additional export menu.
*
*    IF lo_export_btn_choice IS BOUND.
*      zcl_abap2xlsx_helper=>enh_cl_fpm_list_uibb_renderer_(
*        EXPORTING
*          io_export_btn_choice = lo_export_btn_choice
*          io_view              = io_view
*      ).
*    ENDIF.
*
*ENDENHANCEMENT.
*$*$-End:   (1)---------------------------------------------------------------------------------$*$*
  ENDMETHOD.


  METHOD excel_download.
    CHECK: is_abap2xlsx_installed( ) EQ abap_true.
    CALL METHOD ('ZCL_ABAP2XLSX_HELPER_INT')=>('EXCEL_DOWNLOAD')
*    CALL METHOD zcl_abap2xlsx_helper_int=>excel_download(
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


  METHOD MESSAGE.
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
