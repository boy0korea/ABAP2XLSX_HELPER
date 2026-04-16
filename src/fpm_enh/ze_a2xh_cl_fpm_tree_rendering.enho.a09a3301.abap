"Name: \TY:CL_FPM_TREE_RENDERING\ME:RENDER_TOOLBAR_C_TABLE\SE:END\EI
ENHANCEMENT 0 ZE_A2XH_CL_FPM_TREE_RENDERING.
* additional export menu.
    FIELD-SYMBOLS: <lo_export_btn_choice_z> TYPE REF TO cl_wd_toolbar_btn_choice.

*    ASSIGN lo_export_btn_choice_new TO <lo_export_btn_choice_z>.
    ASSIGN ('LO_EXPORT_BTN_CHOICE_NEW') TO <lo_export_btn_choice_z>.
    IF <lo_export_btn_choice_z> IS NOT ASSIGNED.
      " old version
      ASSIGN ('LO_EXPORT_BTN_CHOICE') TO <lo_export_btn_choice_z>.
    ENDIF.

    IF <lo_export_btn_choice_z> IS ASSIGNED AND <lo_export_btn_choice_z> IS BOUND.
      zcl_a2xh_fpm_enh=>enh_cl_fpm_list_uibb_renderer2(
        EXPORTING
          iv_export_format     = ls_settings-export_format
          io_export_btn_choice = <lo_export_btn_choice_z>
          iv_export_action     = lc_export_action
      ).
    ENDIF.

ENDENHANCEMENT.
