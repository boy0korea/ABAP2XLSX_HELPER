METHOD open_popup .

  wd_assist->mo_event_data = io_event_data.
  wd_assist->mo_comp_usage = io_comp_usage.



  DATA lo_window_manager TYPE REF TO if_wd_window_manager.
  DATA lo_api_component  TYPE REF TO if_wd_component.
  DATA lo_window         TYPE REF TO if_wd_window.
  DATA lt_buttons        TYPE wdr_popup_button_list.
  DATA ls_canc_action    TYPE wdr_popup_button_action.

  lo_api_component           = wd_this->wd_get_api( ).
  lo_window_manager          = lo_api_component->get_window_manager( ).
*   create the cancel icon, but without any action handler
  ls_canc_action-action_name = '*'.
*   Simple example, see docu of method create_and_open_popup for details
  lt_buttons                 = lo_window_manager->get_buttons_ok(
*      default_button       = if_wd_window=>co_button_ok
   ).

  lo_window                  = lo_window_manager->create_and_open_popup(
      window_name          = 'W_MAIN'
*      title                =
      message_type         = if_wd_window=>co_msg_type_none
      message_display_mode = if_wd_window=>co_msg_display_mode_selected
*      is_resizable         = ABAP_TRUE
      buttons              = lt_buttons
      cancel_action        = ls_canc_action
  ).


  SET HANDLER wd_assist->on_close FOR lo_window.

ENDMETHOD.

method WDDOAPPLICATIONSTATECHANGE .
endmethod.

method WDDOBEFORENAVIGATION .
endmethod.

method WDDOEXIT .
endmethod.

method WDDOINIT .
endmethod.

method WDDOPOSTPROCESSING .
endmethod.

