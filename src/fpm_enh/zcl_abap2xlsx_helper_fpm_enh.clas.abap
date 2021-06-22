CLASS zcl_abap2xlsx_helper_fpm_enh DEFINITION
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

    CLASS-DATA gv_list_uibb_export_on TYPE flag VALUE abap_true ##NO_TEXT.

    CLASS-METHODS enh_cl_fpm_list_uibb_assist_at
      IMPORTING
        !iv_format       TYPE fpmgb_export_format
        !irt_result_data TYPE REF TO data
        !it_p13n_column  TYPE if_fpm_list_settings_variant=>ty_t_o_column
        !it_field_usage  TYPE fpmgb_t_fieldusage .
    CLASS-METHODS enh_cl_fpm_list_uibb_renderer_
      IMPORTING
        !io_export_btn_choice TYPE REF TO cl_wd_toolbar_btn_choice
        !io_view              TYPE REF TO if_wd_view .
  PROTECTED SECTION.

    CLASS-METHODS readme .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ABAP2XLSX_HELPER_FPM_ENH IMPLEMENTATION.


  METHOD enh_cl_fpm_list_uibb_assist_at.
    DATA: lt_field2      TYPE tt_field,
          lt_field       TYPE tt_field,
          ls_field       TYPE ts_field,
          ls_field_usage TYPE fpmgb_s_fieldusage,
          lo_p13n_column TYPE REF TO if_fpm_list_settings_column,
          lv_column_name TYPE string.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    CHECK: zcl_abap2xlsx_helper=>is_abap2xlsx_installed( ) EQ abap_true,
           gv_list_uibb_export_on EQ abap_true.

    IF iv_format EQ 'ZA2X'.
      ASSIGN irt_result_data->* TO <lt_data>.

      zcl_abap2xlsx_helper=>get_fieldcatalog(
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

      zcl_abap2xlsx_helper=>excel_download(
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
*      zcl_abap2xlsx_helper_fpm_enh=>enh_cl_fpm_list_uibb_assist_at(
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

    CHECK: zcl_abap2xlsx_helper=>is_abap2xlsx_installed( ) EQ abap_true,
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
        id           = `ZMNUAI_FPM_EXPORT_ZA2X`             "#EC NOTEXT
        view         = io_view
        on_action    = 'DISPATCH_EXPORT'  " lif_renderer_constants=>cs_table_action-export
        text         = 'Excel by abap2xlsx'
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
*      zcl_abap2xlsx_helper_fpm_enh=>enh_cl_fpm_list_uibb_renderer_(
*        EXPORTING
*          io_export_btn_choice = lo_export_btn_choice
*          io_view              = io_view
*      ).
*    ENDIF.
*
*ENDENHANCEMENT.
*$*$-End:   (1)---------------------------------------------------------------------------------$*$*
  ENDMETHOD.


  METHOD readme.
* https://github.com/boy0korea/ABAP2XLSX_HELPER
  ENDMETHOD.
ENDCLASS.
