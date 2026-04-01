METHOD close_popup .
  wd_this->wd_get_api( )->get_embedding_window( )->close( ).

ENDMETHOD.

METHOD do_init .
  DATA: lo_popup TYPE REF TO if_wd_window.

  lo_popup = wd_this->wd_get_api( )->get_embedding_window( ).
  IF lo_popup IS NOT INITIAL AND lo_popup->is_modal( ) EQ abap_true.
    lo_popup->set_window_title( 'upload' ).
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


*  io_view->request_focus_on_view_elem( io_view->get_element( 'FILE_UPLOAD' ) ).

  CAST cl_wd_html_fragment( io_view->get_element( 'HTML_FRAGMENT' ) )->set_html(
    `<img src="/sap/public/bc/ur/nw5/1x1.gif" style="display: none;" ` &&
    `onload="` &&
    `var el = this;` &&
    `while( !el.classList.contains('lsPopupWindow') ) {` &&                 " popup div
    `  el = el.parentElement;` &&
    `}` &&
    `if( !el.style.zIndex ) return;` &&
    `var file = el.querySelector('input[type=file]');` &&                   " file input box
    `var ok = el.querySelector('.urPWFooterBottomLine .lsButton');` &&      " ok button
    `file.oninput = function() { UCF_JsUtil.delayedCall( 200, ok, 'click' ); };` &&
    `file.click();` &&
    `" />`
  ).

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
  DATA lo_nd_upload TYPE REF TO if_wd_context_node.
  DATA lo_el_upload TYPE REF TO if_wd_context_element.
  DATA ls_upload TYPE wd_this->element_upload.

  lo_nd_upload = wd_context->get_child_node( name = wd_this->wdctx_upload ).
  lo_el_upload = lo_nd_upload->get_element( ).
  lo_el_upload->get_static_attributes(
    IMPORTING
      static_attributes = ls_upload ).



  wd_assist->on_ok(
    EXPORTING
      iv_file_name    = ls_upload-filename
      iv_file_content = ls_upload-content
  ).

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

