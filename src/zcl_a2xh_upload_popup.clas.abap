class ZCL_A2XH_UPLOAD_POPUP definition
  public
  inheriting from CL_WD_COMPONENT_ASSISTANCE
  create public .

public section.

  data MO_PARAM type ref to IF_FPM_PARAMETER .

  class-methods OPEN_POPUP
    importing
      !IO_PARAM type ref to IF_FPM_PARAMETER .
  methods ON_OK
    importing
      !IV_EXCEL type XSTRING
      !IV_FILENAME type STRING .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_A2XH_UPLOAD_POPUP IMPLEMENTATION.


  METHOD on_ok.
    DATA: lo_fpm      TYPE REF TO if_fpm,
          lv_event_id TYPE fpm_event_id,
          lo_event    TYPE REF TO cl_fpm_event.

    lo_fpm = cl_fpm=>get_instance( ).
    CHECK: lo_fpm IS NOT INITIAL.

    mo_param->get_value(
      EXPORTING
        iv_key   = 'IV_CALLBACK_EVENT_ID'
      IMPORTING
        ev_value = lv_event_id
    ).
    CHECK: lv_event_id IS NOT INITIAL.

    CREATE OBJECT lo_event
      EXPORTING
        iv_event_id = lv_event_id         " This defines the ID of the FPM Event
*       iv_is_validating    = iv_is_validating    " Defines, whether checks need to be performed or not
*       iv_is_transactional = iv_is_transactional " Defines, whether IF_FPM_TRANSACTION is to be processed
*       iv_adapts_context   = iv_adapts_context   " Event changes the adaptation context
*       iv_framework_event  = iv_framework_event  " Event is raised by FPM and not by user or appl. code
*       is_source_uibb      = is_source_uibb      " Source UIBB of Event
*       io_event_data       = io_event_data       " Data for processing
      .

    lo_event->mo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_EXCEL'
        iv_value = iv_excel
    ).

    lo_event->mo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_FILENAME'
        iv_value = iv_filename
    ).

    lo_fpm->raise_event( lo_event ).
  ENDMETHOD.


  METHOD open_popup.
    DATA: lo_comp_usage TYPE REF TO if_wd_component_usage,
          lo_wd_comp    TYPE REF TO ziwci_a2xh_upload_popup.

    cl_wdr_runtime_services=>get_component_usage(
      EXPORTING
        component            = wdr_task=>application->component
        used_component_name  = 'ZA2XH_UPLOAD_POPUP'
        component_usage_name = 'ZA2XH_UPLOAD_POPUP'
        create_component     = abap_true
        do_create            = abap_true
      RECEIVING
        component_usage      = lo_comp_usage
    ).

    lo_wd_comp ?= lo_comp_usage->get_interface_controller( ).
    lo_wd_comp->open_popup(
        io_param = io_param
    ).
  ENDMETHOD.
ENDCLASS.
