"Name: \PR:CL_FPM_LIST_UIBB_RENDERER_ATS=CP\TY:LCL_TABLE_RENDERER\ME:RENDER_STANDARD_TOOLBAR_ITEMS\SE:END\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_LIST_UIBB_RENDE.
* additional export menu.

    IF lo_export_btn_choice IS BOUND.
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_renderer_(
        EXPORTING
          io_export_btn_choice = lo_export_btn_choice
          iv_export_action     = lif_renderer_constants=>cs_table_action-export
      ).
    ENDIF.

ENDENHANCEMENT.
