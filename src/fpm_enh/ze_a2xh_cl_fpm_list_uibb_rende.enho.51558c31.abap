"Name: \PR:CL_FPM_LIST_UIBB_RENDERER_ATS=CP\TY:LCL_TABLE_RENDERER\ME:RENDER_STANDARD_TOOLBAR_ITEMS\SE:END\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_LIST_UIBB_RENDE.
* additional export menu.
    FIELD-SYMBOLS: <lo_export_btn_choice_z> TYPE REF TO cl_wd_toolbar_btn_choice.

*    ASSIGN lo_export_btn_choice_new TO <lo_export_btn_choice_z>.
    ASSIGN ('LO_EXPORT_BTN_CHOICE_NEW') TO <lo_export_btn_choice_z>.
    IF <lo_export_btn_choice_z> IS NOT ASSIGNED.
      " old version
      ASSIGN ('LO_EXPORT_BTN_CHOICE') TO <lo_export_btn_choice_z>.
    ENDIF.

    IF <lo_export_btn_choice_z> IS ASSIGNED AND <lo_export_btn_choice_z> IS BOUND.
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_renderer_(
        EXPORTING
          io_export_btn_choice = <lo_export_btn_choice_z>
          iv_export_action     = lif_renderer_constants=>cs_table_action-export
      ).
    ENDIF.

ENDENHANCEMENT.
