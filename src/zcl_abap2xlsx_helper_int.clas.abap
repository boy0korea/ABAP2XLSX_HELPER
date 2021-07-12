class ZCL_ABAP2XLSX_HELPER_INT definition
  public
  create public .

public section.

  class-methods EXCEL_DOWNLOAD
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods EXCEL_EMAIL
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_SUBJECT type CLIKE optional
      !IV_SENDER type CLIKE optional
      !IT_RECEIVER type STRINGTAB optional
      !IV_FILENAME type CLIKE optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L' .
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
      !IT_DDIC_OBJECT type DD_X031L_TABLE optional
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD optional
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
      !IV_AUTO_COLUMN_WIDTH type FLAG default ABAP_TRUE
      !IV_DEFAULT_DESCR type C default 'L'
    exporting
      !EV_EXCEL type XSTRING
      !EV_ERROR_TEXT type STRING .
  class-methods CONVERT_JSON_TO_EXCEL
    importing
      !IV_DATA_JSON type STRING
      !IT_DDIC_OBJECT type DD_X031L_TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD
      !IV_SHEET_TITLE type CLIKE optional
      !IV_IMAGE_XSTRING type XSTRING optional
      !IV_ADD_FIXEDVALUE_SHEET type FLAG default ABAP_TRUE
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
  class-methods GET_XSTRING_FROM_SMW0
    importing
      !IV_SMW0 type WWWDATA-OBJID
    returning
      value(RV_XSTRING) type XSTRING .
protected section.

  class-methods START_DOWNLOAD
    importing
      !IV_EXCEL type XSTRING
      !IV_FILENAME type CLIKE optional .
  class-methods START_UPLOAD
    exporting
      !EV_EXCEL type XSTRING .
  class-methods DO_DRM_ENCODE
    changing
      !CV_EXCEL type XSTRING .
  class-methods DO_DRM_DECODE
    changing
      !CV_EXCEL type XSTRING .
  class-methods ADD_FIXEDVALUE_SHEET
    importing
      !IT_DATA type STANDARD TABLE
      !IT_FIELD type ZCL_ABAP2XLSX_HELPER=>TT_FIELD
      !IT_FIELD_CATALOG type ZEXCEL_T_FIELDCATALOG
      !IO_EXCEL type ref to ZCL_EXCEL
      !IV_WORKSHEET_INDEX type I default 1
      !IV_HEADER_ROW_INDEX type I default 1
    raising
      ZCX_EXCEL .
  class-methods ADD_IMAGE
    importing
      !IV_IMAGE_XSTRING type XSTRING
      !IV_COL type I
      !IO_EXCEL type ref to ZCL_EXCEL
      !IV_WORKSHEET_INDEX type I default 1
    raising
      ZCX_EXCEL .
  class-methods GET_DDIC_OBJECT
    importing
      !I_DATA type DATA
    returning
      value(RT_DDIC_OBJECT) type DD_X031L_TABLE .
  PRIVATE SECTION.
ENDCLASS.



CLASS ZCL_ABAP2XLSX_HELPER_INT IMPLEMENTATION.


  METHOD add_fixedvalue_sheet.
* http://www.abap2xlsx.org
    DATA: lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lo_worksheet_fv    TYPE REF TO zcl_excel_worksheet,
          lo_data_validation TYPE REF TO zcl_excel_data_validation,
          ls_field_catalog   TYPE zexcel_s_fieldcatalog,
          ls_field           TYPE za2xh_s_fieldcatalog,
          lv_sheet_title_fv  TYPE zexcel_sheet_title,
          lt_comp_view       TYPE abap_component_view_tab,
          ls_comp_view       TYPE abap_simple_componentdescr,
          lt_fixed_value     TYPE ddfixvalues,
          ls_fixed_value     TYPE ddfixvalue,
          lt_ddl             TYPE wdr_context_attr_value_list,
          ls_ddl             TYPE wdr_context_attr_value,
          lv_cell_value	     TYPE zexcel_cell_value,
          lo_style_fv	       TYPE REF TO zcl_excel_style,
          lv_style_fv	       TYPE zexcel_cell_style,
          lv_lines_data      TYPE i,
          lv_lines_ddl       TYPE i,
          lv_index_col       TYPE i.

    lo_worksheet = io_excel->get_worksheet_by_index( iv_worksheet_index ).
    lv_lines_data = lines( it_data ).
    IF lv_lines_data EQ 0.
      lv_lines_data = 1.
    ENDIF.
    lv_lines_data = lv_lines_data + iv_header_row_index.

    lo_style_fv = io_excel->add_new_style( ).
    lo_style_fv->font->color-rgb = zcl_excel_style_color=>c_yellow.
    lv_style_fv = lo_style_fv->get_guid( ).

    CAST cl_abap_structdescr(
      CAST cl_abap_tabledescr(
        cl_abap_tabledescr=>describe_by_data( it_data )
      )->get_table_line_type( )
    )->get_included_view( RECEIVING p_result = lt_comp_view ).
    SORT lt_comp_view BY name.

    LOOP AT it_field_catalog INTO ls_field_catalog.
      lv_index_col = sy-tabix.
      CLEAR: lt_ddl.


      " 1. get from it_field-fxied_values
      READ TABLE it_field INTO ls_field INDEX lv_index_col.
      IF sy-subrc EQ 0.
        lt_ddl = ls_field-fixed_values.
      ENDIF.

      " 2. get from ddic domain fixed value
      IF lt_ddl IS INITIAL.
        READ TABLE lt_comp_view INTO ls_comp_view WITH KEY name = ls_field_catalog-fieldname BINARY SEARCH.
        lt_ddl = zcl_abap2xlsx_helper=>get_ddic_fixed_values( ls_comp_view-type ).
      ENDIF.

      CHECK: lt_ddl IS NOT INITIAL.
      lv_lines_ddl = lines( lt_ddl ) + 1.

      " create fv-sheet
      lv_sheet_title_fv = ls_field_catalog-fieldname.
      lo_worksheet_fv = io_excel->get_worksheet_by_name( lv_sheet_title_fv ).
      IF lo_worksheet_fv IS INITIAL.
        lo_worksheet_fv = io_excel->add_new_worksheet( lv_sheet_title_fv ).
        lo_worksheet_fv->bind_table(
          EXPORTING
            ip_table          = lt_ddl
            it_field_catalog  = VALUE #(
                                  ( fieldname = 'VALUE' position = 1 scrtext_l = 'Value' dynpfld = abap_true )
                                  ( fieldname = 'TEXT' position = 2 scrtext_l = 'Text' dynpfld = abap_true )
                                )
        ).
        lo_worksheet_fv->zif_excel_sheet_protection~protected = lo_worksheet_fv->zif_excel_sheet_protection~c_protected.
        lo_worksheet_fv->zif_excel_sheet_protection~sheet = lo_worksheet_fv->zif_excel_sheet_protection~c_active.
        lo_worksheet_fv->zif_excel_sheet_protection~objects = lo_worksheet_fv->zif_excel_sheet_protection~c_active.
        IF lv_lines_ddl <= 3.
          " If it has 1 or 2 fixed values, hide fv-sheet.
          lo_worksheet_fv->zif_excel_sheet_properties~hidden = lo_worksheet_fv->zif_excel_sheet_properties~c_hidden.
        ENDIF.
      ENDIF.

      " add validation
      lo_data_validation = lo_worksheet->add_new_data_validation( ).
      lo_data_validation->type = zcl_excel_data_validation=>c_type_list.
      lo_data_validation->allowblank = abap_true.
      lo_data_validation->formula1 = lv_sheet_title_fv && '!$A$2:$A$' && lv_lines_ddl.
      lo_data_validation->cell_column = zcl_excel_common=>convert_column2alpha( lv_index_col ).
      lo_data_validation->cell_column_to = lo_data_validation->cell_column.
      lo_data_validation->cell_row = iv_header_row_index + 1.
      lo_data_validation->cell_row_to = lv_lines_data.

      IF iv_header_row_index IS NOT INITIAL AND lo_worksheet_fv->zif_excel_sheet_properties~hidden IS INITIAL.
        " link to fv-sheet @ header
        lo_worksheet->get_cell(
          EXPORTING
            ip_column  = lo_data_validation->cell_column
            ip_row     = iv_header_row_index
          IMPORTING
            ep_value   = lv_cell_value
        ).
        lo_worksheet->set_cell(
          EXPORTING
            ip_column    = lo_data_validation->cell_column
            ip_row       = iv_header_row_index
            ip_value     = lv_cell_value
            ip_style     = lv_style_fv
            ip_hyperlink = zcl_excel_hyperlink=>create_internal_link( lv_sheet_title_fv && '!A1' )
        ).
      ENDIF.

    ENDLOOP.

  ENDMETHOD.


  METHOD add_image.
* http://www.abap2xlsx.org
    DATA: lo_worksheet    TYPE REF TO zcl_excel_worksheet,
          lo_drawing      TYPE REF TO zcl_excel_drawing,
          lv_image_type   TYPE string,
          lv_image_width  TYPE i,
          lv_image_height TYPE i.


    cl_fxs_image_info=>determine_info(
      EXPORTING
        iv_data     = iv_image_xstring
      IMPORTING
        ev_mimetype = lv_image_type
        ev_xres     = lv_image_width
        ev_yres     = lv_image_height
    ).
    CASE lv_image_type.
      WHEN cl_fxs_mime_types=>co_image_bitmap.
        lv_image_type = 'BMP'.
      WHEN cl_fxs_mime_types=>co_image_png.
        lv_image_type = 'PNG'.
      WHEN cl_fxs_mime_types=>co_image_gif.
        lv_image_type = 'GIF'.
      WHEN cl_fxs_mime_types=>co_image_tiff.
        lv_image_type = 'TIF'.
      WHEN cl_fxs_mime_types=>co_image_jpeg.
        lv_image_type = 'JPG'.
      WHEN OTHERS.
        " not supported.
        RETURN.
    ENDCASE.

    lo_worksheet = io_excel->get_worksheet_by_index( iv_worksheet_index ).
    lo_drawing = io_excel->add_new_drawing( ).

    lo_drawing->set_media(
      EXPORTING
        ip_media      = iv_image_xstring
        ip_media_type = lv_image_type
        ip_width      = lv_image_width
        ip_height     = lv_image_height
    ).
    lo_drawing->set_position(
      EXPORTING
        ip_from_row = 1
        ip_from_col = zcl_excel_common=>convert_column2alpha( iv_col )
    ).
    lo_worksheet->add_drawing( lo_drawing ).
  ENDMETHOD.


  METHOD convert_abap_to_excel.
* http://www.abap2xlsx.org
    DATA: lo_excel           TYPE REF TO zcl_excel,
          lo_writer          TYPE REF TO zif_excel_writer,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          ls_table_settings  TYPE zexcel_s_table_settings,
          lt_field_catalog   TYPE zexcel_t_fieldcatalog,
          lt_field_catalog2  TYPE zexcel_t_fieldcatalog,
          ls_field_catalog   TYPE zexcel_s_fieldcatalog,
          ls_field           TYPE za2xh_s_fieldcatalog,
          lo_zcx_excel       TYPE REF TO zcx_excel,
          lv_sheet_title     TYPE zexcel_sheet_title,
          lt_ddic_object     TYPE dd_x031l_table,
          ls_ddic_object     TYPE x031l,
          ls_ddic_object_ref TYPE x031l,
          lv_amount_external TYPE bapicurr-bapicurr,
          lr_data            TYPE REF TO data,
          lv_conversion      TYPE string,
          lv_local_ts        TYPE timestamp,
          lv_alpha_out       TYPE string,
          lv_index_col       TYPE i,
          lv_index           TYPE i.
    FIELD-SYMBOLS: <ls_data>     TYPE data,
                   <lv_data>     TYPE data,
                   <lv_data_ref> TYPE data.
    CLEAR ev_error_text.

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

          IF it_ddic_object IS NOT INITIAL.
            LOOP AT it_ddic_object INTO ls_ddic_object.
              READ TABLE lt_field_catalog2 INTO ls_field_catalog WITH KEY fieldname = ls_ddic_object-fieldname BINARY SEARCH.
              IF sy-subrc EQ 0.
                lv_index = sy-tabix.
                ls_field_catalog-abap_type = ls_ddic_object-exid.
                MODIFY lt_field_catalog2 FROM ls_field_catalog INDEX lv_index TRANSPORTING abap_type.
              ENDIF.
            ENDLOOP.
          ENDIF.

          LOOP AT it_field INTO ls_field.
            lv_index = sy-tabix.
            READ TABLE lt_field_catalog2 INTO ls_field_catalog WITH KEY fieldname = ls_field-fieldname BINARY SEARCH.
            CHECK: sy-subrc EQ 0.
            ls_field_catalog-position = lv_index.
            ls_field_catalog-dynpfld = abap_true.
            IF ls_field-label_text IS NOT INITIAL.
              ls_field_catalog-scrtext_s = ls_field_catalog-scrtext_m = ls_field_catalog-scrtext_l = ls_field-label_text.
            ENDIF.
            APPEND ls_field_catalog TO lt_field_catalog.
          ENDLOOP.
        ENDIF.
        DELETE lt_field_catalog WHERE dynpfld NE abap_true.


**********************************************************************
        lo_worksheet->bind_table( ip_table          = it_data
                                  it_field_catalog  = lt_field_catalog
                                  is_table_settings = ls_table_settings ).
**********************************************************************

        " conversion exit.
        CREATE DATA lr_data LIKE LINE OF it_data.
        ASSIGN lr_data->* TO <ls_data>.
        LOOP AT lt_field_catalog INTO ls_field_catalog.
          lv_index_col = sy-tabix.
          ASSIGN COMPONENT ls_field_catalog-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
          DESCRIBE FIELD <lv_data> EDIT MASK lv_conversion.
          IF lv_conversion EQ '==TSTLC' OR
             lv_conversion EQ '==TSTPS'.
            LOOP AT it_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_field_catalog-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              IF <lv_data> IS NOT INITIAL.
                IF lv_conversion EQ '==TSTLC'.
                  PERFORM convert_to_local_time IN PROGRAM saplsdc_cnv USING <lv_data> CHANGING lv_local_ts.
                ELSE.
                  lv_local_ts = <lv_data>.
                ENDIF.
                lo_worksheet->set_cell(
                  EXPORTING
                    ip_column    = lv_index_col
                    ip_row       = lv_index
                    ip_value     = lv_local_ts
                    ip_abap_type = cl_abap_typedescr=>typekind_char
                ).
              ENDIF.
            ENDLOOP.
          ELSEIF lv_conversion EQ '==ALPHA'.
            LOOP AT it_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_field_catalog-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              IF <lv_data> IS NOT INITIAL.
                CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
                  EXPORTING
                    input  = <lv_data>
                  IMPORTING
                    output = lv_alpha_out.
                lo_worksheet->set_cell(
                  EXPORTING
                    ip_column    = lv_index_col
                    ip_row       = lv_index
                    ip_value     = lv_alpha_out
                    ip_abap_type = cl_abap_typedescr=>typekind_char
                ).
              ENDIF.
            ENDLOOP.
          ENDIF.
        ENDLOOP.

        " currency
        IF it_ddic_object IS NOT INITIAL.
          lt_ddic_object = it_ddic_object.
        ELSE.
          lt_ddic_object = get_ddic_object( it_data ).
        ENDIF.
        LOOP AT lt_ddic_object INTO ls_ddic_object WHERE reffield IS NOT INITIAL.
          READ TABLE lt_field_catalog TRANSPORTING NO FIELDS WITH KEY fieldname = ls_ddic_object-fieldname.
          CHECK: sy-subrc EQ 0.
          lv_index_col = sy-tabix.
          READ TABLE lt_ddic_object INTO ls_ddic_object_ref WITH KEY fieldname = ls_ddic_object-reffield dtyp = 'CUKY'.
          CHECK: sy-subrc EQ 0.
          LOOP AT it_data ASSIGNING <ls_data>.
            lv_index = sy-tabix + 1.
            ASSIGN COMPONENT ls_ddic_object-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
            ASSIGN COMPONENT ls_ddic_object_ref-fieldname OF STRUCTURE <ls_data> TO <lv_data_ref>.
            IF <lv_data> IS NOT INITIAL AND
               <lv_data_ref> IS NOT INITIAL AND
               <lv_data_ref> <> 'USD' AND
               <lv_data_ref> <> 'EUR'.
              CALL FUNCTION 'BAPI_CURRENCY_CONV_TO_EXTERNAL'
                EXPORTING
                  currency        = CONV tcurc-waers( <lv_data_ref> )
                  amount_internal = <lv_data>
                IMPORTING
                  amount_external = lv_amount_external.
              lo_worksheet->set_cell(
                EXPORTING
                  ip_column    = lv_index_col
                  ip_row       = lv_index
                  ip_value     = lv_amount_external
                  ip_abap_type = cl_abap_typedescr=>typekind_packed
              ).
            ENDIF.
          ENDLOOP.
        ENDLOOP.

        " auto column width
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

        " add fixed value sheet
        IF iv_add_fixedvalue_sheet EQ abap_true.
          add_fixedvalue_sheet(
            EXPORTING
              it_data             = it_data
              it_field            = it_field
              it_field_catalog    = lt_field_catalog
              io_excel            = lo_excel
          ).
          lo_excel->set_active_sheet_index( 1 ).
        ENDIF.

        " add image
        IF iv_image_xstring IS NOT INITIAL.
          add_image(
            EXPORTING
              iv_image_xstring   = iv_image_xstring
              iv_col             = lines( lt_field_catalog ) + 1
              io_excel           = lo_excel
          ).
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


    do_drm_decode(
      CHANGING
        cv_excel = ev_excel
    ).

  ENDMETHOD.


  METHOD convert_excel_to_abap.
* http://www.abap2xlsx.org
    DATA: lo_excel                    TYPE REF TO zcl_excel,
          lo_reader                   TYPE REF TO zif_excel_reader,
          lo_worksheet                TYPE REF TO zcl_excel_worksheet,
          lv_excel                    TYPE xstring,
          lt_field                    TYPE za2xh_t_fieldcatalog,
          ls_field                    TYPE za2xh_s_fieldcatalog,
          lv_highest_column           TYPE int4,
          lv_highest_row              TYPE int4,
          lv_column                   TYPE int4,
          lv_col_str                  TYPE zexcel_cell_column_alpha,
          lv_row                      TYPE int4,
          lv_value                    TYPE zexcel_cell_value,
          lv_date                     TYPE datum,
          lv_time                     TYPE uzeit,
          lv_char_row                 TYPE string,
          lv_style_guid               TYPE zexcel_cell_style,
          ls_stylemapping             TYPE zexcel_s_stylemapping,
          lt_ddic_object              TYPE dd_x031l_table,
          ls_ddic_object              TYPE x031l,
          ls_ddic_object_ref          TYPE x031l,
          lv_ddic_object_has_currency TYPE flag,
          lv_amount_external          TYPE bapicurr-bapicurr,
          lo_root                     TYPE REF TO cx_root,
          lo_zcx_excel                TYPE REF TO zcx_excel,
          lv_conversion               TYPE string,
          lv_local_ts                 TYPE timestamp,
          lv_index_col                TYPE i,
          lv_index                    TYPE i.
    FIELD-SYMBOLS: <ls_data>     TYPE data,
                   <lv_data>     TYPE data,
                   <lv_data_ref> TYPE data.

    CLEAR: ev_error_text, et_data[].

    CHECK: iv_excel IS NOT INITIAL.
    lv_excel = iv_excel.
    do_drm_decode(
      CHANGING
        cv_excel = lv_excel
    ).


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
        lo_excel = lo_reader->load( lv_excel  ). "Load data into reader
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
              IF sy-subrc EQ 0.
                ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
                IF sy-subrc EQ 0.
                  <lv_data> = lv_value.
                ENDIF.
              ENDIF.
            ENDIF.
            lv_column = lv_column + 1.
          ENDWHILE.

          " delete empty line
          IF <ls_data> IS INITIAL.
            DELETE et_data INDEX lines( et_data ).
          ENDIF.
          lv_row = lv_row + 1.
        ENDWHILE.




        " conversion exit.
        LOOP AT lt_field INTO ls_field.
          ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
          DESCRIBE FIELD <lv_data> EDIT MASK lv_conversion.
          IF lv_conversion EQ '==TSTLC'.
            LOOP AT et_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              IF <lv_data> IS NOT INITIAL.
                PERFORM convert_to_utc_time IN PROGRAM saplsdc_cnv USING <lv_data> CHANGING <lv_data>.
              ENDIF.
            ENDLOOP.
          ELSEIF lv_conversion EQ '==ALPHA'.
            LOOP AT et_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_field-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
                EXPORTING
                  input  = <lv_data>
                IMPORTING
                  output = <lv_data>.
            ENDLOOP.
          ENDIF.
        ENDLOOP.

        lt_ddic_object = get_ddic_object( et_data ).

        " currency
        LOOP AT lt_ddic_object INTO ls_ddic_object WHERE reffield IS NOT INITIAL.
          READ TABLE lt_ddic_object INTO ls_ddic_object_ref WITH KEY fieldname = ls_ddic_object-reffield dtyp = 'CUKY'.
          CHECK: sy-subrc EQ 0.
          LOOP AT et_data ASSIGNING <ls_data>.
            lv_index = sy-tabix + 1.
            ASSIGN COMPONENT ls_ddic_object-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
            IF <lv_data> IS NOT INITIAL.
              ASSIGN COMPONENT ls_ddic_object_ref-fieldname OF STRUCTURE <ls_data> TO <lv_data_ref>.
              IF <lv_data_ref> IS NOT INITIAL AND
                 <lv_data_ref> <> 'USD' AND
                 <lv_data_ref> <> 'EUR'.

                lv_amount_external = <lv_data>.
                CALL FUNCTION 'BAPI_CURRENCY_CONV_TO_INTERNAL'
                  EXPORTING
                    currency             = CONV tcurc-waers( <lv_data_ref> )
                    amount_external      = lv_amount_external
                    max_number_of_digits = 23
                  IMPORTING
                    amount_internal      = <lv_data>.
              ENDIF.
            ENDIF.
          ENDLOOP.
        ENDLOOP.


      CATCH cx_root INTO lo_root.
        ev_error_text = lo_root->get_text( ).
        IF lv_row IS NOT INITIAL.
          lv_char_row = lv_row.
          CONDENSE lv_char_row NO-GAPS.
          ev_error_text = ev_error_text && '(Col:' && lv_col_str && ',Row:' && lv_char_row && ')'.
        ENDIF.
    ENDTRY.



  ENDMETHOD.


  METHOD convert_json_to_excel.
* http://www.abap2xlsx.org
    DATA: ls_field    TYPE za2xh_s_fieldcatalog,
          lt_comp     TYPE abap_component_tab,
          ls_comp     TYPE abap_componentdescr,
          lo_tab_type TYPE REF TO cl_abap_tabledescr,
          lr_data     TYPE REF TO data.
    FIELD-SYMBOLS: <lt_data> TYPE table.

    " json to table
    LOOP AT it_field INTO ls_field.
      ls_comp-name = ls_field-fieldname.
      ls_comp-type = cl_abap_elemdescr=>get_string( ).
      APPEND ls_comp TO lt_comp.
    ENDLOOP.
    lo_tab_type = cl_abap_tabledescr=>create( cl_abap_structdescr=>create( lt_comp ) ).
    CREATE DATA lr_data TYPE HANDLE lo_tab_type.
    ASSIGN lr_data->* TO <lt_data>.

    /ui2/cl_json=>deserialize(
      EXPORTING
        json             = iv_data_json
      CHANGING
        data             = <lt_data>
    ).

    " table to excel
    convert_abap_to_excel(
      EXPORTING
        it_data                 = <lt_data>
        it_ddic_object          = it_ddic_object
        it_field                = it_field
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr
      IMPORTING
        ev_excel                = ev_excel
        ev_error_text           = ev_error_text
    ).
  ENDMETHOD.


  METHOD do_drm_decode.
* if you need to DRM decode. write code here.
  ENDMETHOD.


  METHOD do_drm_encode.
* if you need to DRM encode. write code here.
  ENDMETHOD.


  METHOD excel_download.
* http://www.abap2xlsx.org
    convert_abap_to_excel(
      EXPORTING
        it_data                 = it_data
        it_field                = it_field
        iv_sheet_title          = iv_sheet_title
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr
      IMPORTING
        ev_excel                = ev_excel
        ev_error_text           = ev_error_text
    ).

    IF ev_error_text IS NOT INITIAL.
      zcl_abap2xlsx_helper=>message( iv_error_text = ev_error_text ).
      RETURN.
    ENDIF.

    start_download(
      EXPORTING
        iv_excel    = ev_excel
        iv_filename = iv_filename
    ).

  ENDMETHOD.


  METHOD excel_email.
* http://www.abap2xlsx.org
    DATA: lt_receiver   TYPE TABLE OF string,
          lo_event_data TYPE REF TO if_fpm_parameter.

    IF it_receiver IS NOT INITIAL.
      lt_receiver = it_receiver.
    ELSE.
      lt_receiver = zcl_za2xh_email_popup=>get_default_receiver( ).
    ENDIF.

    lo_event_data = NEW cl_fpm_parameter( ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IT_RECEIVER'
        iv_value = lt_receiver
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IT_DATA'
        iv_value = it_data
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IT_FIELD'
        iv_value = it_field
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_SUBJECT'
        iv_value = CONV string( iv_subject )
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_SENDER'
        iv_value = CONV string( iv_sender )
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_FILENAME'
        iv_value = CONV string( iv_filename )
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_SHEET_TITLE'
        iv_value = CONV string( iv_sheet_title )
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_IMAGE_XSTRING'
        iv_value = iv_image_xstring
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_ADD_FIXEDVALUE_SHEET'
        iv_value = iv_add_fixedvalue_sheet
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_AUTO_COLUMN_WIDTH'
        iv_value = iv_auto_column_width
    ).

    lo_event_data->set_value(
      EXPORTING
        iv_key   = 'IV_DEFAULT_DESCR'
        iv_value = iv_default_descr
    ).


    IF wdr_task=>application IS NOT INITIAL.
      " WD or FPM
      zcl_za2xh_email_popup=>open_popup( io_event_data = lo_event_data ).
    ELSE.
      " GUI
      CALL FUNCTION 'ZA2XH_EMAIL_POPUP_GUI'
        EXPORTING
          io_event_data = lo_event_data.
    ENDIF.


    IF 1 EQ 2.
      " ok click on popup
      NEW zcl_za2xh_email_popup( )->on_ok( ).
    ENDIF.

  ENDMETHOD.


  METHOD excel_upload.
* http://www.abap2xlsx.org
    DATA: lv_excel TYPE xstring.
    FIELD-SYMBOLS: <lv_excel> TYPE xstring.


    IF iv_excel IS NOT INITIAL.
      ASSIGN iv_excel TO <lv_excel>.
    ELSE.
      start_upload(
        IMPORTING
          ev_excel = lv_excel
      ).
      ASSIGN lv_excel TO <lv_excel>.
    ENDIF.

    convert_excel_to_abap(
      EXPORTING
        iv_excel      = <lv_excel>
        it_field      = it_field
        iv_begin_row  = iv_begin_row
        iv_sheet_no   = iv_sheet_no
      IMPORTING
        et_data       = et_data
        ev_error_text = ev_error_text
    ).

    IF ev_error_text IS NOT INITIAL.
      zcl_abap2xlsx_helper=>message( iv_error_text = ev_error_text ).
    ENDIF.

  ENDMETHOD.


  METHOD get_fieldcatalog.
* http://www.abap2xlsx.org
    DATA: lt_fc    TYPE zexcel_t_fieldcatalog,
          ls_fc    TYPE zexcel_s_fieldcatalog,
          ls_field TYPE za2xh_s_fieldcatalog.

    lt_fc = zcl_excel_common=>get_fieldcatalog( ip_table = it_data ).
    DELETE lt_fc WHERE dynpfld NE abap_true.

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


  METHOD get_xstring_from_smw0.
    DATA: lo_drawing TYPE REF TO zcl_excel_drawing.

    CREATE OBJECT lo_drawing.
    lo_drawing->set_media_www(
      EXPORTING
        ip_key    = VALUE #( relid = 'MI' objid = iv_smw0 )
        ip_width  = 0
        ip_height = 0
    ).
    rv_xstring = lo_drawing->get_media( ).
  ENDMETHOD.


  METHOD start_download.
    DATA: lv_filename_string   TYPE string,
          lv_filename_path     TYPE string,
          lv_filename_fullpath TYPE string,
          lv_bin_filesize      TYPE i,
          lt_temptable         TYPE w3mimetabtype.


    IF iv_filename IS NOT INITIAL.
      lv_filename_string = iv_filename.
    ELSE.
      lv_filename_string = zcl_abap2xlsx_helper=>default_excel_filename( ).
    ENDIF.

    IF wdr_task=>application IS NOT INITIAL.
      " WD or FPM
      CALL METHOD cl_wd_runtime_services=>attach_file_to_response
        EXPORTING
          i_filename  = lv_filename_string
          i_content   = iv_excel
          i_mime_type = 'xlsx'.
    ELSE.
      " GUI
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
    DATA: lt_file_table	TYPE filetable,
          ls_file_table TYPE file_table,
          lv_rc	        TYPE i,
          lv_filename   TYPE string,
          lv_filelength TYPE i,
          lt_temptable  TYPE w3mimetabtype.

    IF wdr_task=>application IS NOT INITIAL.
      " WD or FPM
      RETURN.
    ELSE.
      " GUI
      cl_gui_frontend_services=>file_open_dialog(
        EXPORTING
*         window_title            = window_title      " Title Of File Open Dialog
          default_extension       = 'xlsx' " Default Extension
*         default_filename        = default_filename  " Default File Name
          file_filter             = 'excel (*.xlsx)|*.xlsx|'       " File Extension Filter String
*         with_encoding           = with_encoding     " File Encoding
*         initial_directory       = initial_directory " Initial Directory
          multiselection          = abap_false    " Multiple selections poss.
        CHANGING
          file_table              = lt_file_table        " Table Holding Selected Files
          rc                      = lv_rc                " Return Code, Number of Files or -1 If Error Occurred
*         user_action             = user_action       " User Action (See Class Constants ACTION_OK, ACTION_CANCEL)
*         file_encoding           = file_encoding
        EXCEPTIONS
          file_open_dialog_failed = 1                 " "Open File" dialog failed
          cntl_error              = 2                 " Control error
          error_no_gui            = 3                 " No GUI available
          not_supported_by_gui    = 4                 " GUI does not support this
          OTHERS                  = 5
      ).
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      READ TABLE lt_file_table INTO ls_file_table INDEX 1.
      CHECK: sy-subrc EQ 0.
      lv_filename = ls_file_table-filename.

      cl_gui_frontend_services=>gui_upload(
        EXPORTING
          filename                = lv_filename              " Name of file
          filetype                = 'BIN'              " File Type (ASCII, Binary)
*         has_field_separator     = space              " Columns Separated by Tabs in Case of ASCII Upload
*         header_length           = 0                  " Length of Header for Binary Data
*         read_by_line            = 'X'                " File Written Line-By-Line to the Internal Table
*         dat_mode                = space              " Numeric and date fields are in DAT format in WS_DOWNLOAD
*         codepage                = codepage           " Character Representation for Output
*         ignore_cerr             = abap_true          " Ignore character set conversion errors?
*         replacement             = '#'                " Replacement Character for Non-Convertible Characters
*         virus_scan_profile      = virus_scan_profile " Virus Scan Profile
        IMPORTING
          filelength              = lv_filelength         " File Length
*         header                  = header             " File Header in Case of Binary Upload
        CHANGING
          data_tab                = lt_temptable           " Transfer table for file contents
*         isscanperformed         = space              " File already scanned
        EXCEPTIONS
          file_open_error         = 1                  " File does not exist and cannot be opened
          file_read_error         = 2                  " Error when reading file
          no_batch                = 3                  " Cannot execute front-end function in background
          gui_refuse_filetransfer = 4                  " Incorrect front end or error on front end
          invalid_type            = 5                  " Incorrect parameter FILETYPE
          no_authority            = 6                  " No upload authorization
          unknown_error           = 7                  " Unknown error
          bad_data_format         = 8                  " Cannot Interpret Data in File
          header_not_allowed      = 9                  " Invalid header
          separator_not_allowed   = 10                 " Invalid separator
          header_too_long         = 11                 " Header information currently restricted to 1023 bytes
          unknown_dp_error        = 12                 " Error when calling data provider
          access_denied           = 13                 " Access to file denied.
          dp_out_of_memory        = 14                 " Not enough memory in data provider
          disk_full               = 15                 " Storage medium is full.
          dp_timeout              = 16                 " Data provider timeout
          not_supported_by_gui    = 17                 " GUI does not support this
          error_no_gui            = 18                 " GUI not available
          OTHERS                  = 19
      ).
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
        EXPORTING
          input_length = lv_filelength
*         first_line   = 0
*         last_line    = 0
        IMPORTING
          buffer       = ev_excel
        TABLES
          binary_tab   = lt_temptable
        EXCEPTIONS
          failed       = 1
          OTHERS       = 2.
      IF sy-subrc <> 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.
    ENDIF.
  ENDMETHOD.


  METHOD get_ddic_object.
    DATA: lo_type_desc   TYPE REF TO cl_abap_typedescr,
          lo_table_desc  TYPE REF TO cl_abap_tabledescr,
          lo_struct_desc TYPE REF TO cl_abap_structdescr.

    lo_type_desc = cl_abap_typedescr=>describe_by_data( i_data ).
    CASE lo_type_desc->kind.
      WHEN cl_abap_typedescr=>kind_table.
        lo_table_desc ?= lo_type_desc.
        lo_type_desc = lo_table_desc->get_table_line_type( ).
        IF lo_type_desc->kind EQ cl_abap_typedescr=>kind_struct.
          lo_struct_desc ?= lo_type_desc.
        ENDIF.
      WHEN cl_abap_typedescr=>kind_struct.
        lo_struct_desc ?= lo_type_desc.
    ENDCASE.

    IF lo_struct_desc IS NOT INITIAL.
      lo_struct_desc->get_ddic_object(
        RECEIVING
          p_object     = rt_ddic_object
        EXCEPTIONS
          not_found    = 1        " Type could not be found
          no_ddic_type = 2        " Typ is not a dictionary type
          OTHERS       = 3
      ).
    ENDIF.

  ENDMETHOD.
ENDCLASS.
