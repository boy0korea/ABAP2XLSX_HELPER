METHOD close_popup .
  wd_this->wd_get_api( )->get_embedding_window( )->close( ).

ENDMETHOD.

METHOD do_init .
  DATA: lo_popup TYPE REF TO if_wd_window.

  lo_popup = wd_this->wd_get_api( )->get_embedding_window( ).
  IF lo_popup IS NOT INITIAL AND lo_popup->is_modal( ) EQ abap_true.
    lo_popup->set_window_title( 'send email' ).
    lo_popup->set_close_button( abap_true ).
    lo_popup->set_close_in_any_case( abap_false ).
    lo_popup->set_button_kind( if_wd_window=>co_buttons_okcancel ).
    lo_popup->set_default_button( if_wd_window=>co_button_none ).
    lo_popup->set_on_close_action(
      EXPORTING
        view               = io_view
        action_name        = 'POPUP_BUTTON'
    ).
    lo_popup->subscribe_to_button_event(
      EXPORTING
        button            = if_wd_window=>co_button_ok
        action_name       = 'POPUP_BUTTON'
        action_view       = io_view
    ).
    lo_popup->subscribe_to_button_event(
      EXPORTING
        button            = if_wd_window=>co_button_cancel
        action_name       = 'POPUP_BUTTON'
        action_view       = io_view
    ).
  ENDIF.


  io_view->request_focus_on_view_elem( io_view->get_element( 'INPUT_TOKENIZER' ) ).

  DATA: lt_string TYPE TABLE OF string,
        lv_string TYPE string.
  wd_assist->mo_event_data->get_value(
    EXPORTING
      iv_key   = 'IT_RECEIVER'
    IMPORTING
      ev_value = lt_string
  ).

  DATA lo_nd_receiver TYPE REF TO if_wd_context_node.
  DATA lt_receiver TYPE wd_this->elements_receiver.
  DATA ls_receiver TYPE wd_this->element_receiver.

  IF lt_string IS NOT INITIAL.
    LOOP AT lt_string INTO lv_string.
      ls_receiver-email = lv_string.
      APPEND ls_receiver TO lt_receiver.
    ENDLOOP.
    APPEND INITIAL LINE TO lt_receiver.

    lo_nd_receiver = wd_context->get_child_node( name = wd_this->wdctx_receiver ).
    lo_nd_receiver->bind_table( new_items = lt_receiver set_initial_elements = abap_true ).
    lo_nd_receiver->set_lead_selection_index( lines( lt_receiver ) ).
  ENDIF.



ENDMETHOD.

METHOD onactioninput_tokenizer .
  DATA: lo_node     TYPE REF TO if_wd_context_node,
        lt_receiver TYPE TABLE OF wd_this->element_receiver,
        ls_receiver TYPE wd_this->element_receiver,
        lt_string   TYPE TABLE OF string,
        lv_string   TYPE string.

  lo_node = context_element->get_node( ).

  context_element->get_static_attributes(
    IMPORTING
      static_attributes = ls_receiver ).

  CASE wdevent->name.
    WHEN 'ON_ENTER'
      OR 'ON_PASTE'.
      lt_string = wd_assist->split_email_string( ls_receiver-email ).
      LOOP AT lt_string INTO lv_string.
        ls_receiver-email = lv_string.
        APPEND ls_receiver TO lt_receiver.
      ENDLOOP.

      lo_node->bind_elements(
        EXPORTING
          new_items            = lt_receiver " List of Elements or Model Data
          set_initial_elements = abap_false " If TRUE, Set Initial Elements Otherwise Add
          index                = context_element->get_index( )     " Index of Context Element
      ).

      context_element->set_static_attributes_null( ).
      lo_node->move_element(
        EXPORTING
          from = context_element->get_index( )
          to   = lo_node->get_element_count( )
      ).


    WHEN 'ON_EDIT'.
      lo_node->remove_element( lo_node->get_lead_selection( ) ).
      lo_node->set_lead_selection( context_element ).


    WHEN 'ON_DELETE'.
      lo_node->remove_element( context_element ).


  ENDCASE.
ENDMETHOD.

METHOD onactionpopup_button .

  CASE wdevent->name.
    WHEN 'ON_OK'
      OR 'ON_YES'.
      " OK button
      on_ok( ).
    WHEN OTHERS.
      " other button (Cancel, Close ...)
      close_popup( ).
  ENDCASE.
ENDMETHOD.

METHOD on_ok .
  wd_assist->on_ok( ).
  close_popup( ).
ENDMETHOD.

method WDDOAFTERACTION .

endmethod.

method WDDOBEFOREACTION .
*  data lo_api_controller type ref to if_wd_view_controller.
*  data lo_action         type ref to if_wd_action.

*  lo_api_controller = wd_this->wd_get_api( ).
*  lo_action = lo_api_controller->get_current_action( ).

*  if lo_action is bound.
*    case lo_action->name.
*      when '...'.

*    endcase.
*  endif.

endmethod.

method WDDOEXIT .

endmethod.

method WDDOINIT .

endmethod.

METHOD wddomodifyview .
  IF first_time EQ abap_true.
    do_init( view ).
  ENDIF.

ENDMETHOD.

method WDDOONCONTEXTMENU .

endmethod.

