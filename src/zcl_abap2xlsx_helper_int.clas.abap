class ZCL_ABAP2XLSX_HELPER_INT definition
  public
  create public .

public section.

  class-methods CLASS_CONSTRUCTOR .
  class-methods EXCEL_DOWNLOAD
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods EXCEL_UPLOAD
    importing
      !IV_EXCEL type XSTRING optional
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_BEGIN_ROW type INT4 default 2
      !IV_SHEET_NO type INT1 default 1
    exporting
      !ET_DATA type STANDARD TABLE
      !EV_ERROR_TEXT type STRING .
  class-methods GET_FIELDCATALOG
    importing
      !IT_DATA type STANDARD TABLE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !ET_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD .
  class-methods CONVERT_ABAP_TO_EXCEL
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods CONVERT_EXCEL_TO_ABAP
    importing
      !IV_EXCEL type XSTRING
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_BEGIN_ROW type INT4 default 2
      !IV_SHEET_NO type INT1 default 1
    exporting
      !ET_DATA type STANDARD TABLE
      !EV_ERROR_TEXT type STRING .
  PROTECTED SECTION.

    CLASS-METHODS start_download
      IMPORTING
        !iv_excel    TYPE xstring
        !iv_filename TYPE clike OPTIONAL .
    CLASS-METHODS start_upload .
    CLASS-METHODS do_drm_encode
      CHANGING
        !cv_excel TYPE xstring .
    CLASS-METHODS do_drm_decode
      CHANGING
        !cv_excel TYPE xstring .
    CLASS-METHODS message
      IMPORTING
        !iv_error_text TYPE clike .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ABAP2XLSX_HELPER_INT IMPLEMENTATION.


  METHOD class_constructor.
  ENDMETHOD.


  METHOD convert_abap_to_excel.
* http://www.abap2xlsx.org

    DATA: lo_excel             TYPE REF TO zcl_excel,
          lo_writer            TYPE REF TO zif_excel_writer,
          lo_worksheet         TYPE REF TO zcl_excel_worksheet,
          ls_table_settings    TYPE zexcel_s_table_settings,
          lt_field_catalog     TYPE zexcel_t_fieldcatalog,
          lt_field_catalog2    TYPE zexcel_t_fieldcatalog,
          ls_field_catalog     TYPE zexcel_s_fieldcatalog,
          ls_field             TYPE zcl_abap2xlsx_helper=>ts_field,
          lo_zcx_excel         TYPE REF TO zcx_excel,
          lv_sheet_title       TYPE zexcel_sheet_title,
          lv_filename_string   TYPE string,
          lv_filename_path     TYPE string,
          lv_filename_fullpath TYPE string,
          lv_bin_filesize      TYPE i,
          lt_temptable         TYPE w3mimetabtype,
          lv_index             TYPE i.
    CLEAR ev_error_text.

* TODO: fixed value 자동 탭 추가 기능.

    TRY.

        " Creates active sheet
        CREATE OBJECT lo_excel.

        " Get active sheet
        IF iv_sheet_title IS NOT INITIAL.
          lv_sheet_title = iv_sheet_title.
        ELSE.
          lv_sheet_title = 'Export'.
        ENDIF.
        lo_worksheet = lo_excel->get_active_worksheet( ).
        lo_worksheet->set_title( ip_title = lv_sheet_title ).

        " table settings
        ls_table_settings-table_style       = zcl_excel_table=>builtinstyle_medium2.
        ls_table_settings-show_row_stripes  = abap_true.
        ls_table_settings-nofilters         = abap_false.

        " field catalog
        lt_field_catalog = zcl_excel_common=>get_fieldcatalog( ip_table = it_data ).
        IF it_field IS NOT INITIAL.
          lt_field_catalog2 = lt_field_catalog.
          SORT lt_field_catalog2 BY fieldname.
          CLEAR: lt_field_catalog.
          LOOP AT it_field INTO ls_field.
            lv_index = sy-tabix.
            READ TABLE lt_field_catalog2 INTO ls_field_catalog WITH KEY fieldname = ls_field-fieldname BINARY SEARCH.
            CHECK: sy-subrc EQ 0.
            ls_field_catalog-position = lv_index.
            IF ls_field-label_text IS NOT INITIAL.
              ls_field_catalog-scrtext_s = ls_field_catalog-scrtext_m = ls_field_catalog-scrtext_l = ls_field-label_text.
            ENDIF.
            APPEND ls_field_catalog TO lt_field_catalog.
          ENDLOOP.
        ENDIF.


        lo_worksheet->bind_table( ip_table          = it_data
                                  it_field_catalog  = lt_field_catalog
                                  is_table_settings = ls_table_settings ).

        IF iv_auto_column_width EQ abap_true.
          lv_index = lines( lt_field_catalog ).
          DO lv_index TIMES.
            lo_worksheet->set_column_width(
              EXPORTING
                ip_column         = sy-index
                ip_width_autosize = abap_true
            ).
          ENDDO.
        ENDIF.

        "freeze column headers when scrolling
        lo_worksheet->freeze_panes( ip_num_rows = 1 ).

        " Create output
        CREATE OBJECT lo_writer TYPE zcl_excel_writer_2007.
        ev_excel = lo_writer->write_file( lo_excel ).


      CATCH zcx_excel INTO lo_zcx_excel.    " Exceptions for ABAP2XLSX
        ev_error_text = lo_zcx_excel->error.
        RETURN.
    ENDTRY.


  ENDMETHOD.


  METHOD convert_excel_to_abap.
* http://www.abap2xlsx.org

    DATA: lo_excel          TYPE REF TO zcl_excel,
          lo_reader         TYPE REF TO zif_excel_reader,
          lo_worksheet      TYPE REF TO zcl_excel_worksheet,
          lt_field          TYPE zcl_abap2xlsx_helper=>tt_field,
          ls_field          TYPE zcl_abap2xlsx_helper=>ts_field,
          lv_highest_column TYPE int4,
          lv_highest_row    TYPE int4,
          lv_column         TYPE int4,
          lv_col_str        TYPE zexcel_cell_column_alpha,
          lv_row            TYPE int4,
          lv_value          TYPE zexcel_cell_value,
          lv_date           TYPE datum,
          lv_time           TYPE uzeit,
          lv_char_row       TYPE string,
          lv_style_guid     TYPE zexcel_cell_style,
          ls_stylemapping   TYPE zexcel_s_stylemapping,
          lo_root           TYPE REF TO cx_root,
          lo_zcx_excel      TYPE REF TO zcx_excel.
    FIELD-SYMBOLS: <ls_data> TYPE data,
                   <lv_data> TYPE data.

    CLEAR: ev_error_text, et_data[].

    TRY.
        IF it_field IS NOT INITIAL.
          lt_field = it_field.
        ELSE.
          get_fieldcatalog(
            EXPORTING
              it_data          = et_data
            IMPORTING
              et_field         = lt_field
          ).
        ENDIF.

        CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
        lo_excel = lo_reader->load( iv_excel  ). "Load data into reader
        lo_excel->set_active_sheet_index( iv_sheet_no ).
        lo_worksheet = lo_excel->get_active_worksheet( ).

        lv_highest_column = lo_worksheet->get_highest_column( ).
        lv_highest_row    = lo_worksheet->get_highest_row( ).
        lv_row = iv_begin_row.

        WHILE lv_row <= lv_highest_row.
          APPEND INITIAL LINE TO et_data ASSIGNING <ls_data>.
          lv_column = 1.

          WHILE lv_column <= lv_highest_column.
            lv_col_str = zcl_excel_common=>convert_column2alpha( lv_column ).
            lo_worksheet->get_cell(
              EXPORTING
                ip_column = lv_col_str
                ip_row    = lv_row
              IMPORTING
                ep_value  = lv_value
                ep_guid   = lv_style_guid ).

            IF lv_style_guid IS NOT INITIAL AND lv_value IS NOT INITIAL.
              " Read style attributes
              ls_stylemapping = lo_excel->get_style_to_guid( lv_style_guid ).
              CASE ls_stylemapping-complete_style-number_format-format_code.
                WHEN zcl_excel_style_number_format=>c_format_date_ddmmyyyy
                  OR zcl_excel_style_number_format=>c_format_date_ddmmyyyydot
                  OR zcl_excel_style_number_format=>c_format_date_dmminus
                  OR zcl_excel_style_number_format=>c_format_date_dmyminus
                  OR zcl_excel_style_number_format=>c_format_date_dmyslash
                  OR zcl_excel_style_number_format=>c_format_date_myminus
                  OR zcl_excel_style_number_format=>c_format_date_std
                  OR zcl_excel_style_number_format=>c_format_date_xlsx14
                  OR zcl_excel_style_number_format=>c_format_date_xlsx15
                  OR zcl_excel_style_number_format=>c_format_date_xlsx16
                  OR zcl_excel_style_number_format=>c_format_date_xlsx17
                  OR zcl_excel_style_number_format=>c_format_date_xlsx22
                  OR zcl_excel_style_number_format=>c_format_date_xlsx45
                  OR zcl_excel_style_number_format=>c_format_date_xlsx46
                  OR zcl_excel_style_number_format=>c_format_date_xlsx47
                  OR zcl_excel_style_number_format=>c_format_date_yymmdd
                  OR zcl_excel_style_number_format=>c_format_date_yymmddminus
                  OR zcl_excel_style_number_format=>c_format_date_yymmddslash
                  OR zcl_excel_style_number_format=>c_format_date_yyyymmdd
                  OR zcl_excel_style_number_format=>c_format_date_yyyymmddminus
                  OR zcl_excel_style_number_format=>c_format_date_yyyymmddslash.
                  " Convert excel date to ABAP date
                  lv_date = zcl_excel_common=>excel_string_to_date( lv_value ).
                  lv_value = lv_date.
                WHEN zcl_excel_style_number_format=>c_format_date_time1
                  OR zcl_excel_style_number_format=>c_format_date_time2
                  OR zcl_excel_style_number_format=>c_format_date_time3
                  OR zcl_excel_style_number_format=>c_format_date_time4
                  OR zcl_excel_style_number_format=>c_format_date_time5
                  OR zcl_excel_style_number_format=>c_format_date_time6
                  OR zcl_excel_style_number_format=>c_format_date_time7
                  OR zcl_excel_style_number_format=>c_format_date_time8
                  OR 'h:mm:ss;@'.
                  " Convert excel time to ABAP time
                  lv_time = zcl_excel_common=>excel_string_to_time( lv_value ).
                  lv_value = lv_time.
              ENDCASE.
            ENDIF.

            CONDENSE lv_value.
            IF lv_value IS NOT INITIAL.
*              ASSIGN COMPONENT lv_column OF STRUCTURE <ls_data> TO <lv_data>.
              READ TABLE lt_field INTO ls_field INDEX lv_column.
              CHECK: sy-subrc EQ 0.
              ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              CHECK: sy-subrc EQ 0.
              <lv_data> = lv_value.
            ENDIF.
            lv_column = lv_column + 1.
          ENDWHILE.

          IF <ls_data> IS INITIAL.
            DELETE et_data INDEX lines( et_data ).
          ENDIF.
          lv_row    = lv_row + 1.
        ENDWHILE.

      CATCH cx_root INTO lo_root.
        ev_error_text = lo_root->get_text( ).
        IF lv_row IS NOT INITIAL.
          lv_char_row = lv_row.
          CONDENSE lv_char_row NO-GAPS.
          ev_error_text = ev_error_text && '(Col:' && lv_col_str && ',Row:' && lv_char_row && ')'.
        ENDIF.
    ENDTRY.



  ENDMETHOD.


  METHOD do_drm_decode.
  ENDMETHOD.


  METHOD do_drm_encode.
  ENDMETHOD.


  METHOD excel_download.
* http://www.abap2xlsx.org
    convert_abap_to_excel(
      EXPORTING
        it_data              = it_data
        it_field             = it_field
        iv_sheet_title       = iv_sheet_title
        iv_auto_column_width = iv_auto_column_width
        iv_default_descr     = iv_default_descr
      IMPORTING
        ev_excel             = ev_excel
        ev_error_text        = ev_error_text
    ).

    IF ev_error_text IS NOT INITIAL.
      message( iv_error_text = ev_error_text ).
      RETURN.
    ENDIF.

    do_drm_decode(
      CHANGING
        cv_excel = ev_excel
    ).

    start_download(
      EXPORTING
        iv_excel    = ev_excel
        iv_filename = iv_filename
    ).

  ENDMETHOD.


  METHOD excel_upload.
    DATA: lv_excel TYPE xstring.

    lv_excel = iv_excel.

    IF lv_excel IS INITIAL.
      start_upload( ).
    ENDIF.

    CHECK: lv_excel IS NOT INITIAL.

    do_drm_decode(
      CHANGING
        cv_excel = lv_excel
    ).

    convert_excel_to_abap(
      EXPORTING
        iv_excel      = lv_excel
        it_field      = it_field
        iv_begin_row  = iv_begin_row
        iv_sheet_no   = iv_sheet_no
      IMPORTING
        et_data       = et_data
        ev_error_text = ev_error_text
    ).

    IF ev_error_text IS NOT INITIAL.
      message( iv_error_text = ev_error_text ).
    ENDIF.

  ENDMETHOD.


  METHOD get_fieldcatalog.
* http://www.abap2xlsx.org
    DATA: lt_fc    TYPE zexcel_t_fieldcatalog,
          ls_fc    TYPE zexcel_s_fieldcatalog,
          ls_field TYPE zcl_abap2xlsx_helper=>ts_field.

    lt_fc = zcl_excel_common=>get_fieldcatalog( ip_table = it_data ).
    DELETE lt_fc WHERE dynpfld = abap_false.

    LOOP AT lt_fc INTO ls_fc.
      ls_field-fieldname = ls_fc-fieldname.
      CASE iv_default_descr.
        WHEN 'L'.
          ls_field-label_text = ls_fc-scrtext_l.
        WHEN 'M'.
          ls_field-label_text = ls_fc-scrtext_m.
        WHEN 'S'.
          ls_field-label_text = ls_fc-scrtext_s.
        WHEN OTHERS.
          CLEAR: ls_field-label_text.
      ENDCASE.
      APPEND ls_field TO et_field.
    ENDLOOP.
  ENDMETHOD.


  METHOD message.
    CHECK: iv_error_text IS NOT INITIAL.

    IF wdr_task=>application IS NOT INITIAL.
      " WD or FPM
      wdr_task=>application->component->if_wd_controller~get_message_manager( )->report_error_message(
        EXPORTING
          message_text = iv_error_text
      ).
    ELSE.
      " GUI
      MESSAGE iv_error_text TYPE 'S' DISPLAY LIKE 'E'.
    ENDIF.

  ENDMETHOD.


  METHOD start_download.
    DATA: lv_filename_string   TYPE string,
          lv_filename_path     TYPE string,
          lv_filename_fullpath TYPE string,
          lv_bin_filesize      TYPE i,
          lt_temptable         TYPE w3mimetabtype,
          lv_index             TYPE i.


    lv_filename_string = iv_filename.
    IF iv_filename IS INITIAL.
      lv_filename_string = |export_{ sy-datum }_{ sy-uzeit }.xlsx|.
    ENDIF.
    IF wdr_task=>application_name IS NOT INITIAL.
      CALL METHOD cl_wd_runtime_services=>attach_file_to_response
        EXPORTING
          i_filename  = lv_filename_string
          i_content   = iv_excel
          i_mime_type = 'xlsx'.
    ELSE.
      CALL FUNCTION 'SCMS_XSTRING_TO_BINARY'
        EXPORTING
          buffer        = iv_excel
        IMPORTING
          output_length = lv_bin_filesize
        TABLES
          binary_tab    = lt_temptable.
      cl_gui_frontend_services=>file_save_dialog(
        EXPORTING
          default_file_name         = lv_filename_string " DEFAULT FILE NAME
          file_filter               = '*.xlsx'       " file type filter table
        CHANGING
          filename                  = lv_filename_string          " file name to save
          path                      = lv_filename_path              " path to file
          fullpath                  = lv_filename_fullpath          " path + file name
        EXCEPTIONS
          OTHERS                    = 5
      ).
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.
      IF lv_filename_fullpath IS NOT INITIAL.
        cl_gui_frontend_services=>gui_download(
           EXPORTING
             bin_filesize = lv_bin_filesize
             filename     = lv_filename_fullpath
             filetype     = 'BIN'
           CHANGING
             data_tab     = lt_temptable
           EXCEPTIONS
             file_write_error          = 1
             no_batch                  = 2
             gui_refuse_filetransfer   = 3
             invalid_type              = 4
             no_authority              = 5
             unknown_error             = 6
             header_not_allowed        = 7
             separator_not_allowed     = 8
             filesize_not_allowed      = 9
             header_too_long           = 10
             dp_error_create           = 11
             dp_error_send             = 12
             dp_error_write            = 13
             unknown_dp_error          = 14
             access_denied             = 15
             dp_out_of_memory          = 16
             disk_full                 = 17
             dp_timeout                = 18
             file_not_found            = 19
             dataprovider_exception    = 20
             control_flush_error       = 21
             not_supported_by_gui      = 22
             error_no_gui              = 23
             OTHERS                    = 24
        ).
        IF sy-subrc <> 0.
          MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
                     WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
        ENDIF.
      ENDIF.
    ENDIF.

  ENDMETHOD.


  METHOD start_upload.
  ENDMETHOD.
ENDCLASS.
