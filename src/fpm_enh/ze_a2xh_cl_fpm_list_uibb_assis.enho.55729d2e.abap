"Name: \PR:CL_FPM_LIST_UIBB_ASSIST_ATS===CP\TY:LCL_EXPORT_ACTION\ME:EXECUTE\SE:BEGIN\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_LIST_UIBB_ASSIS.
* additional export menu.
    DATA: zlrt_result_data TYPE REF TO data.

    IF me->mv_format CP 'Z*'.
      me->get_result_data(
        EXPORTING
          iv_data_only        = abap_true
        IMPORTING
          ert_result_data     = zlrt_result_data
      ).
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_assist_at(
        EXPORTING
          iv_format       = me->mv_format
          irt_result_data = zlrt_result_data
          it_p13n_column  = me->mo_list_uibb_assist->mo_personalization_api->get_current_variant( )->get_columns( )
          it_field_usage  = me->mo_list_uibb_assist->mt_field_usage
          iv_from_comp    = me->mo_list_uibb_assist->mv_component_name
          io_c_table      = cast CL_FPM_LIST_UIBB_RENDERER_ATS( me->mo_list_uibb_assist->mo_render )->get_table_ui_element( )
      ).
      RETURN.
    ENDIF.

ENDENHANCEMENT.
