"Name: \TY:CL_FPM_TREE_RENDERING\ME:RENDER_TOOLBAR_C_TABLE\SE:END\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_TREE_RENDERING.
* additional export menu.

    IF lo_export_btn_choice IS BOUND.
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_renderer_(
        EXPORTING
          io_export_btn_choice = lo_export_btn_choice
          iv_export_action     = lc_export_action
      ).
    ENDIF.

ENDENHANCEMENT.
