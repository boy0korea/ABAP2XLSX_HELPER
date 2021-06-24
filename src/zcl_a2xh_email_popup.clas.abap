class ZCL_A2XH_EMAIL_POPUP definition
  public
  inheriting from CL_WD_COMPONENT_ASSISTANCE
  create public .

public section.

  class-methods OPEN_POPUP
    importing
      !IO_PARAM type ref to IF_FPM_PARAMETER .
  methods ON_OK
    importing
      !IO_PARAM type ref to IF_FPM_PARAMETER .
protected section.
private section.
ENDCLASS.



CLASS ZCL_A2XH_EMAIL_POPUP IMPLEMENTATION.


  METHOD on_ok.
    DATA: lr_data     TYPE REF TO data,
          lt_receiver TYPE TABLE OF string.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    io_param->get_value(
      EXPORTING
        iv_key   = 'IT_DATA'
      IMPORTING
*        ev_value = ev_value
        er_value = lr_data
    ).
    ASSIGN lr_data->* TO <lt_data>.

    io_param->get_value(
      EXPORTING
        iv_key   = 'IT_RECEIVER'
      IMPORTING
        ev_value = lt_receiver
    ).


    CALL FUNCTION 'ZA2XH_EMAIL'
      EXPORTING
        it_data     = <lt_data>
*       it_field    = it_field
*       iv_filename = iv_filename
*       iv_subject  = iv_subject
*       iv_sender   = iv_sender
        it_receiver = lt_receiver.
  ENDMETHOD.


  METHOD open_popup.
    DATA: lo_comp_usage TYPE REF TO if_wd_component_usage,
          lo_wd_comp    TYPE REF TO ziwci_a2xh_email_popup.

    cl_wdr_runtime_services=>get_component_usage(
      EXPORTING
        component            = wdr_task=>application->component
        used_component_name  = 'ZA2XH_EMAIL_POPUP'
        component_usage_name = 'ZA2XH_EMAIL_POPUP'
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
