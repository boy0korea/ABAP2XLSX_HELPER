FUNCTION za2xh_email_rfc.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     VALUE(IV_DATA_JSON) TYPE  STRING
*"     VALUE(IT_DDIC_OBJECT) TYPE  DD_X031L_TABLE
*"     VALUE(IT_FIELD) TYPE  ZA2XH_T_FIELDCATALOG
*"     VALUE(IV_SUBJECT) TYPE  STRING OPTIONAL
*"     VALUE(IV_SENDER) TYPE  STRING OPTIONAL
*"     VALUE(IT_RECEIVER) TYPE  STRINGTAB
*"     VALUE(IV_FILENAME) TYPE  STRING OPTIONAL
*"     VALUE(IV_SHEET_TITLE) TYPE  STRING OPTIONAL
*"     VALUE(IV_ADD_FIXEDVALUE_SHEET) TYPE  FLAG DEFAULT ABAP_TRUE
*"     VALUE(IV_AUTO_COLUMN_WIDTH) TYPE  FLAG DEFAULT ABAP_TRUE
*"     VALUE(IV_DEFAULT_DESCR) TYPE  CHAR1 DEFAULT 'L'
*"----------------------------------------------------------------------
  DATA: lv_excel       TYPE xstring,
        lo_bcs         TYPE REF TO cl_bcs,
        lv_subject     TYPE string,
        lv_string      TYPE string,
        lv_email       TYPE ad_smtpadr,
        lv_filename    TYPE sood-objdes,
        lv_filesize    TYPE sood-objlen,
        lt_filecontent TYPE solix_tab,
        lo_document    TYPE REF TO cl_document_bcs.
  FIELD-SYMBOLS: <lt_table> TYPE table.

  CHECK: iv_data_json IS NOT INITIAL AND it_field IS NOT INITIAL.

  zcl_abap2xlsx_helper=>convert_json_to_excel(
    EXPORTING
      iv_data_json            = iv_data_json
      it_ddic_object          = it_ddic_object
      it_field                = it_field
      iv_sheet_title          = iv_sheet_title
      iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
      iv_auto_column_width    = iv_auto_column_width
      iv_default_descr        = iv_default_descr
    IMPORTING
      ev_excel                = lv_excel
*      ev_error_text           = ev_error_text
  ).
  CHECK: lv_excel IS NOT INITIAL.

  " filename
  IF iv_filename IS NOT INITIAL.
    lv_filename = iv_filename.
  ELSE.
    lv_filename = zcl_abap2xlsx_helper=>default_excel_filename( ).
  ENDIF.

  " subject
  IF iv_subject IS NOT INITIAL.
    lv_subject = iv_subject.
  ELSE.
    lv_subject = lv_filename.
  ENDIF.


  TRY.
      lo_bcs = cl_bcs=>create_persistent( ).

      cl_document_bcs=>create_document(
        EXPORTING
          i_type         = 'RAW'         " Code for Document Class
          i_subject      = CONV #( lv_subject )      " Short Description of Contents
*          i_length       = i_length       " Size of Document Content
*          i_language     = space          " Language in Which Document Is Created
*          i_importance   = i_importance   " Document priority
*          i_sensitivity  = i_sensitivity  " Object: Sensitivity (private, functional, ...)
          i_text         = VALUE #( ( line =  'attached' ) )         " Content (Text-Like)
*          i_hex          = i_hex          " Content (Binary)
*          i_header       = i_header       " Objcont and Objhead as Table Type
*          i_sender       = i_sender       " BCS: Represents a BAS Address
*          iv_vsi_profile = iv_vsi_profile " Virus Scan Profile
        RECEIVING
          result         = lo_document         " Wrapper Class for Office Documents
      ).

      " sender
      IF iv_sender IS NOT INITIAL.
        lv_email = iv_sender.
        lo_bcs->set_sender( cl_cam_address_bcs=>create_internet_address( lv_email ) ).
      ENDIF.

      " receiver
      LOOP AT it_receiver INTO lv_string.
        lv_email = lv_string.
        lo_bcs->add_recipient( cl_cam_address_bcs=>create_internet_address( lv_email ) ).
      ENDLOOP.


      " attachment
      lv_filesize = xstrlen( lv_excel ).
      lt_filecontent = cl_bcs_convert=>xstring_to_solix( lv_excel ).

      lo_document->add_attachment(
        EXPORTING
          i_attachment_type     = 'xls'    " Document Class for Attachment
          i_attachment_subject  = lv_filename " Attachment Title
          i_attachment_size     = lv_filesize    " Size of Document Content
*          i_attachment_language = space                " Language in Which Attachment Is Created
*          i_att_content_text    = i_att_content_text   " Content (Text-Like)
          i_att_content_hex     = lt_filecontent    " Content (Binary)
*          i_attachment_header   = i_attachment_header  " Attachment Header Data
*          iv_vsi_profile        = iv_vsi_profile       " Virus Scan Profile
      ).

      " send
      lo_bcs->set_document( lo_document ).
      lo_bcs->set_message_subject( lv_subject ).
      lo_bcs->send_without_dialog( ).

      COMMIT WORK.

    CATCH cx_bcs.

  ENDTRY.

ENDFUNCTION.
