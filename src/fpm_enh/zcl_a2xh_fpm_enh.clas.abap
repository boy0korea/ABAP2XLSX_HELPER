class ZCL_A2XH_FPM_ENH definition
  public
  create public .

public section.

  class-data GV_LIST_UIBB_EXPORT_ON type FLAG value ABAP_TRUE ##NO_TEXT.

  class-methods ENH_CL_FPM_LIST_UIBB_ASSIST_AT
    importing
      !IV_FORMAT type FPMGB_EXPORT_FORMAT
      !IRT_RESULT_DATA type ref to DATA
      !IT_P13N_COLUMN type IF_FPM_LIST_SETTINGS_VARIANT=>TY_T_O_COLUMN
      !IT_FIELD_USAGE type FPMGB_T_FIELDUSAGE
      !IV_FROM_COMP type WDY_COMPONENT_NAME
      !IO_C_TABLE type ref to CL_WD_C_TABLE .
  class-methods ENH_CL_FPM_LIST_UIBB_RENDERER_
    importing
      !IO_EXPORT_BTN_CHOICE type ref to CL_WD_TOOLBAR_BTN_CHOICE
      !IV_EXPORT_ACTION type STRING .
  PROTECTED SECTION.

    CLASS-METHODS readme .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_A2XH_FPM_ENH IMPLEMENTATION.


  METHOD enh_cl_fpm_list_uibb_assist_at.
    DATA: lt_field2      TYPE zcl_abap2xlsx_helper=>tt_field,
          lt_field       TYPE zcl_abap2xlsx_helper=>tt_field,
          ls_field       TYPE zcl_abap2xlsx_helper=>ts_field,
          ls_field_usage TYPE fpmgb_s_fieldusage,
          lo_p13n_column TYPE REF TO if_fpm_list_settings_column,
          lv_column_name TYPE string.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    CHECK: gv_list_uibb_export_on EQ abap_true AND
           zcl_abap2xlsx_helper=>is_abap2xlsx_installed( iv_with_message = abap_false ) EQ abap_true.

    ASSIGN irt_result_data->* TO <lt_data>.

    io_c_table->get_data_source( )->get_static_attributes_table(
      IMPORTING
        table = <lt_data>
    ).

    zcl_abap2xlsx_helper=>get_fieldcatalog(
      EXPORTING
        it_data          = <lt_data>
      IMPORTING
        et_field         = lt_field2
    ).

    IF iv_from_comp EQ if_fpm_constants=>gc_components-tree.
      ls_field-fieldname = 'MASTER_COLUMN_TEXT'.
      ls_field-label_text = io_c_table->get_column( id = 'MASTER_COLUMN' )->get_header( )->get_text( ).
      APPEND ls_field TO lt_field.
    ENDIF.

    LOOP AT it_p13n_column INTO lo_p13n_column.
      CHECK: lo_p13n_column->is_visible( ).
      lv_column_name = lo_p13n_column->get_name( ).
      READ TABLE lt_field2 INTO ls_field WITH KEY fieldname = lv_column_name.
      CHECK: sy-subrc EQ 0.

      IF iv_from_comp EQ if_fpm_constants=>gc_components-tree.
        ls_field-label_text = io_c_table->get_column( id = lv_column_name && '_C' )->get_header( )->get_text( ).
      ELSE.
        ls_field-label_text = io_c_table->get_column( id = lv_column_name )->get_header( )->get_text( ).
      ENDIF.

      READ TABLE it_field_usage INTO ls_field_usage WITH KEY name = lv_column_name.
*      ls_field-label_text = ls_field_usage-label_text.
      ls_field-fixed_values = ls_field_usage-fixed_values.
      APPEND ls_field TO lt_field.
    ENDLOOP.

    CASE iv_format.
      WHEN 'ZA2X'.
        zcl_abap2xlsx_helper=>excel_download(
          EXPORTING
            it_data              = <lt_data>
            it_field             = lt_field
        ).
      WHEN 'ZA2E'.
        zcl_abap2xlsx_helper=>excel_email(
          EXPORTING
            it_data                 = <lt_data>
            it_field                = lt_field
        ).
    ENDCASE.


* enhancement 위치:
*Enhanced Development Object    CL_FPM_LIST_UIBB_ASSIST_ATS
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""$"$\SE:(1) Class LCL_EXPORT_ACTION, Method EXECUTE, Start                                                                                                    A
*$*$-Start: (1)---------------------------------------------------------------------------------$*$*
*ENHANCEMENT 1  ZE_ABAP2XLSX_HELPER_LIST_ASSIS.    "active version
** additional export menu.
*    DATA: zlrt_result_data TYPE REF TO data.
*
*    IF me->mv_format CP 'Z*'.
*      me->get_result_data(
*        EXPORTING
*          iv_data_only        = abap_true
*        IMPORTING
*          ert_result_data     = zlrt_result_data
*      ).
*      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_assist_at(
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
    DATA: lo_tab_action TYPE REF TO cl_wd_menu_action_item,
          lo_el         TYPE REF TO cl_abap_elemdescr,
          lt_fv         TYPE ddfixvalues,
          ls_fv         TYPE ddfixvalue.

    CHECK: gv_list_uibb_export_on EQ abap_true AND
           zcl_abap2xlsx_helper=>is_abap2xlsx_installed( iv_with_message = abap_false ) EQ abap_true.

    CHECK: io_export_btn_choice IS BOUND.

*     io_export_btn_choice->remove_choice(
*       EXPORTING
*         id         = 'MNUAI_FPM_EXPORT_PDF'
*     ).

*    lo_tab_action =
*      cl_wd_menu_action_item=>new_menu_action_item(
*        id           = `ZMNUAI_FPM_EXPORT_ZA2X`             "#EC NOTEXT
*        on_action    = 'DISPATCH_EXPORT'  " lif_renderer_constants=>cs_table_action-export
*        text         = 'Excel by abap2xlsx'
*        enabled      = abap_true
*        visible      = abap_true
*    ).
*    DATA(lt_action_parameters) = VALUE wdr_name_value_list(
*      (
*         name  = 'FORMAT'   " lc_export_action_format_param
*         value = 'ZA2X'
*      )
*    ).
*    lo_tab_action->map_on_action( lt_action_parameters ).
*    io_export_btn_choice->add_choice( lo_tab_action ).

    lo_el ?= cl_abap_elemdescr=>describe_by_data( if_fpm_list_types=>cs_export_format-selection_at_runtime ).
    lt_fv = lo_el->get_ddic_fixed_values( ).

    LOOP AT lt_fv INTO ls_fv WHERE low CP 'Z*'.
      lo_tab_action =
        cl_wd_menu_action_item=>new_menu_action_item(
          id           = `MNUAI_FPM_EXPORT_` && ls_fv-low   "#EC NOTEXT
          on_action    = iv_export_action  " lif_renderer_constants=>cs_table_action-export
          text         = CONV string( ls_fv-ddtext )
          enabled      = abap_true
          visible      = abap_true
      ).
      DATA(lt_action_parameters) = VALUE wdr_name_value_list(
        (
           name  = 'FORMAT'   " lc_export_action_format_param
           value = ls_fv-low
        )
      ).
      lo_tab_action->map_on_action( lt_action_parameters ).
      io_export_btn_choice->add_choice( lo_tab_action ).
    ENDLOOP.


* enhancement 위치:
*Enhanced Development Object    CL_FPM_LIST_UIBB_RENDERER_ATS
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""$"$\SE:(1) Class LCL_TABLE_RENDERER, Method RENDER_STANDARD_TOOLBAR_ITEMS, End                                                                               A
*$*$-Start: (1)---------------------------------------------------------------------------------$*$*
*ENHANCEMENT 1  ZE_ABAP2XLSX_HELPER_LIST_RENDE.    "active version
** additional export menu.
*
*    IF lo_export_btn_choice IS BOUND.
*      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_renderer_(
*        EXPORTING
*          io_export_btn_choice = lo_export_btn_choice
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
