"Name: \PR:CL_FPM_TREE_ASSISTANCE========CP\TY:LCL_EXPORT_ACTION\ME:EXECUTE\SE:BEGIN\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_TREE_ASSISTANCE.
* additional export menu.
    DATA: zlrt_result_data TYPE REF TO data,
          zlt_columns TYPE CL_FPM_GUIBB_BASE_RENDER=>T_COLUMN.

    IF me->mv_format CP 'Z*'.
*      zlt_columns = me->mo_tree_uibb_assist->mo_config_data->get_columns( iv_include_secondary_editors = abap_true ).
      zlt_columns = me->mo_tree_uibb_assist->mo_config_data->get_columns( ).
      me->create_result_data(
        EXPORTING
          irt_columns                = ref #( zlt_columns )
          iv_mc_image_src_field_name = me->get_master_col_image_src_field( )
          irt_frontend_data          = me->get_frontend_data( )
        RECEIVING
          rrt_result_data            = zlrt_result_data
      ).
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_assist_at(
        EXPORTING
          iv_format       = me->mv_format
          irt_result_data = zlrt_result_data
          it_p13n_column  = me->mo_tree_uibb_assist->mo_personalization_api->get_current_variant( )->get_columns( )
          it_field_usage  = me->mo_tree_uibb_assist->mt_field_usage
          iv_from_comp    = me->mo_tree_uibb_assist->mv_component_name
          io_c_table      = me->mo_tree_uibb_assist->mo_render_tree->mo_c_table
      ).
      RETURN.
    ENDIF.

ENDENHANCEMENT.
