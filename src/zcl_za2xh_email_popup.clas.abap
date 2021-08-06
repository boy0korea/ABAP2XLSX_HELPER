CLASS zcl_za2xh_email_popup DEFINITION
  PUBLIC
  INHERITING FROM cl_wd_component_assistance
  CREATE PUBLIC .

  PUBLIC SECTION.

    DATA mo_event_data TYPE REF TO if_fpm_parameter .
    DATA mo_comp_usage TYPE REF TO if_wd_component_usage .
    CLASS-DATA gv_wd_comp_id TYPE string READ-ONLY .
    CLASS-DATA go_wd_comp TYPE REF TO ziwci_a2xh_email_popup READ-ONLY .

    CLASS-METHODS class_constructor .
    CLASS-METHODS open_popup
      IMPORTING
        !io_event_data TYPE REF TO if_fpm_parameter OPTIONAL .
    METHODS on_ok .
    METHODS on_close
        FOR EVENT window_closed OF if_wd_window .
    CLASS-METHODS split_email_string
      IMPORTING
        !iv_input       TYPE clike
      RETURNING
        VALUE(rt_email) TYPE stringtab .
    CLASS-METHODS get_default_receiver
      RETURNING
        VALUE(rt_receiver) TYPE stringtab .
  PROTECTED SECTION.
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ZA2XH_EMAIL_POPUP IMPLEMENTATION.


  METHOD class_constructor.
    gv_wd_comp_id = CAST cl_abap_refdescr( cl_abap_typedescr=>describe_by_data( go_wd_comp ) )->get_referenced_type( )->get_relative_name( ).
    REPLACE 'IWCI_' IN gv_wd_comp_id WITH ''.
  ENDMETHOD.


  METHOD get_default_receiver.
    DATA: lv_my_email    TYPE string,
          lt_error_table TYPE TABLE OF rpbenerr.

    CALL FUNCTION 'HR_FBN_GET_USER_EMAIL_ADDRESS'
      EXPORTING
        user_id       = sy-uname
        reaction      = 'N'
      IMPORTING
        email_address = lv_my_email
*       subrc         = subrc         " Return Value, Return Value After ABAP Statements
      TABLES
        error_table   = lt_error_table.   " Benefit structure for error table

    IF lv_my_email IS NOT INITIAL.
      APPEND lv_my_email TO rt_receiver.
    ENDIF.
  ENDMETHOD.


  METHOD on_close.
    mo_comp_usage->delete_component( ).
  ENDMETHOD.


  METHOD on_ok.
    DATA: lr_data                 TYPE REF TO data,
          lt_field                TYPE za2xh_t_fieldcatalog,
          lv_subject              TYPE string,
          lv_sender               TYPE string,
          lt_receiver             TYPE TABLE OF string,
          lv_filename             TYPE string,
          lv_sheet_title          TYPE string,
          lv_image_xstring        TYPE xstring,
          lv_add_fixedvalue_sheet TYPE flag,
          lv_auto_column_width    TYPE flag,
          lv_default_descr        TYPE char1.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IT_DATA'
      IMPORTING
*        ev_value = ev_value
        er_value = lr_data
    ).
    ASSIGN lr_data->* TO <lt_data>.
    CHECK: sy-subrc EQ 0.

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IT_FIELD'
      IMPORTING
        ev_value = lt_field
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_SUBJECT'
      IMPORTING
        ev_value = lv_subject
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_SENDER'
      IMPORTING
        ev_value = lv_sender
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IT_RECEIVER'
      IMPORTING
        ev_value = lt_receiver
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_FILENAME'
      IMPORTING
        ev_value = lv_filename
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_SHEET_TITLE'
      IMPORTING
        ev_value = lv_sheet_title
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_IMAGE_XSTRING'
      IMPORTING
        ev_value = lv_image_xstring
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_ADD_FIXEDVALUE_SHEET'
      IMPORTING
        ev_value = lv_add_fixedvalue_sheet
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_AUTO_COLUMN_WIDTH'
      IMPORTING
        ev_value = lv_auto_column_width
    ).

    mo_event_data->get_value(
      EXPORTING
        iv_key   = 'IV_DEFAULT_DESCR'
      IMPORTING
        ev_value = lv_default_descr
    ).


    CALL FUNCTION 'ZA2XH_EMAIL'
      EXPORTING
        it_data                 = <lt_data>
        it_field                = lt_field
        iv_subject              = lv_subject
        iv_sender               = lv_sender
        it_receiver             = lt_receiver
        iv_filename             = lv_filename
        iv_sheet_title          = lv_sheet_title
        iv_image_xstring        = lv_image_xstring
        iv_add_fixedvalue_sheet = lv_add_fixedvalue_sheet
        iv_auto_column_width    = lv_auto_column_width
        iv_default_descr        = lv_default_descr.

  ENDMETHOD.


  METHOD open_popup.
    DATA: lo_comp_usage TYPE REF TO if_wd_component_usage.

    cl_wdr_runtime_services=>get_component_usage(
      EXPORTING
        component            = wdr_task=>application->component
        used_component_name  = gv_wd_comp_id
        component_usage_name = gv_wd_comp_id
        create_component     = abap_true
        do_create            = abap_true
      RECEIVING
        component_usage      = lo_comp_usage
    ).

    go_wd_comp ?= lo_comp_usage->get_interface_controller( ).

    go_wd_comp->open_popup(
        io_event_data = io_event_data
        io_comp_usage = lo_comp_usage
    ).
  ENDMETHOD.


  METHOD split_email_string.
    DATA: lv_string TYPE string.

    lv_string = iv_input.

    REPLACE ALL OCCURRENCES OF REGEX '[[:space:]]' IN lv_string WITH `;`.
    SPLIT lv_string AT ';' INTO TABLE rt_email.
    DELETE rt_email WHERE table_line IS INITIAL.

  ENDMETHOD.
ENDCLASS.
