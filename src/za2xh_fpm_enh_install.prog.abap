*&---------------------------------------------------------------------*
*& Report ZA2XH_FPM_ENH_INSTALL
*&---------------------------------------------------------------------*
*&
*&---------------------------------------------------------------------*
REPORT za2xh_fpm_enh_install.

PARAMETERS: p_switch TYPE sfw_switchpos AS LISTBOX VISIBLE LENGTH 10.

**********************************************************************
INITIALIZATION.
**********************************************************************
  CALL FUNCTION 'SFW_GET_SWITCH_STATE'
    EXPORTING
      switch_name  = 'ZS_A2XH_FPM_ENH'
    IMPORTING
      switch_state = p_switch.

**********************************************************************
START-OF-SELECTION.
**********************************************************************


  DATA:
    system_settings    TYPE REF TO cl_sfw_system_settings,
    incompatible_bfset TYPE REF TO cx_sfw_incompatible_bfset,
    message_exception  TYPE REF TO cx_sfw_message,
    errors             TYPE swbme_error_tab,
    error_entry        TYPE sfw_wb_checklist_entry,
    bfuncs             TYPE if_sfw_domains=>ty_bfunc_tab.
  FIELD-SYMBOLS:
    <switched_bfunc> TYPE sfw_active_b2,
    <bfunc>          TYPE if_sfw_domains=>ty_bfunc_entry,
    <error>          TYPE swbme_error_entry.

  " Sicherheitsinitiative:
  " - Berechtigungsprüfung auf S_SWITCH
  " - Aktivität 07 (Aktivieren)
  " - Objekttyp 'SFBF' (Business Function)
  AUTHORITY-CHECK OBJECT 'S_SWITCH'
    ID 'ACTVT'   FIELD '07'
    ID 'OBJTYPE' FIELD 'SFBF'
    ID 'OBJNAME' DUMMY.

  IF sy-subrc <> 0.
    MESSAGE e024(sfw).
  ENDIF.

  system_settings = cl_sfw_system_settings=>get_instance( ).

  TRY.
      system_settings->lock( ).

*     Select Business Function Set
      system_settings->select_bfset( 'ZBS_A2XH_FPM_ENH' ).
*     Switch on/off Business Functions

      IF p_switch EQ 'T'.
*           Switch on
        system_settings->switch_bfunc( bfunc = 'ZBF_A2XH_FPM_ENH' on = abap_true ).
      ELSE.
*           Switch off
        system_settings->switch_bfunc( bfunc = 'ZBF_A2XH_FPM_ENH' on = abap_false ).
      ENDIF.

      system_settings->save( ).
      system_settings->activate( ).
    CATCH cx_sfw_message INTO message_exception.
      sy-msgid = message_exception->if_t100_message~t100key-msgid.
      sy-msgno = message_exception->if_t100_message~t100key-msgno.
      sy-msgv1 = message_exception->message_variable_1.
      sy-msgv2 = message_exception->message_variable_2.
      sy-msgv3 = message_exception->message_variable_3.
      sy-msgv4 = message_exception->message_variable_4.
      MESSAGE ID sy-msgid TYPE 'I' NUMBER sy-msgno
        WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
  ENDTRY.

  system_settings->unlock( ).
