CLASS zcl_abap2xlsx_helper_int DEFINITION
  PUBLIC
  CREATE PUBLIC .

  PUBLIC SECTION.

    CLASS-METHODS excel_download
      IMPORTING
        !it_data                 TYPE STANDARD TABLE
        !it_field                TYPE za2xh_t_fieldcatalog OPTIONAL
        !iv_filename             TYPE clike OPTIONAL
        !iv_sheet_title          TYPE clike OPTIONAL
        !iv_image_xstring        TYPE xstring OPTIONAL
        !iv_add_fixedvalue_sheet TYPE flag DEFAULT abap_true
        !iv_auto_column_width    TYPE flag DEFAULT abap_true
        !iv_default_descr        TYPE c DEFAULT 'L'
      EXPORTING
        !ev_excel                TYPE xstring
        !ev_error_text           TYPE string .
    CLASS-METHODS excel_email
      IMPORTING
        !it_data                 TYPE STANDARD TABLE
        !it_field                TYPE za2xh_t_fieldcatalog OPTIONAL
        !iv_subject              TYPE clike OPTIONAL
        !iv_sender               TYPE clike OPTIONAL
        !it_receiver             TYPE stringtab OPTIONAL
        !iv_filename             TYPE clike OPTIONAL
        !iv_sheet_title          TYPE clike OPTIONAL
        !iv_image_xstring        TYPE xstring OPTIONAL
        !iv_add_fixedvalue_sheet TYPE flag DEFAULT abap_true
        !iv_auto_column_width    TYPE flag DEFAULT abap_true
        !iv_default_descr        TYPE c DEFAULT 'L' .
    CLASS-METHODS excel_upload
      IMPORTING
        !iv_excel      TYPE xstring OPTIONAL
        !it_field      TYPE za2xh_t_fieldcatalog OPTIONAL
        !iv_begin_row  TYPE int4 DEFAULT 2
        !iv_sheet_no   TYPE int1 DEFAULT 1
      EXPORTING
        !et_data       TYPE STANDARD TABLE
        !ev_error_text TYPE string
        !et_error_log  TYPE za2xh_t_error_log .
    CLASS-METHODS get_fieldcatalog
      IMPORTING
        !it_data          TYPE STANDARD TABLE
        !iv_default_descr TYPE c DEFAULT 'L'
      EXPORTING
        !et_field         TYPE za2xh_t_fieldcatalog .
    CLASS-METHODS convert_abap_to_excel
      IMPORTING
        !it_data                 TYPE STANDARD TABLE
        !it_ddic_object          TYPE dd_x031l_table OPTIONAL
        !it_field                TYPE za2xh_t_fieldcatalog OPTIONAL
        !iv_sheet_title          TYPE clike OPTIONAL
        !iv_image_xstring        TYPE xstring OPTIONAL
        !iv_add_fixedvalue_sheet TYPE flag DEFAULT abap_true
        !iv_auto_column_width    TYPE flag DEFAULT abap_true
        !iv_default_descr        TYPE c DEFAULT 'L'
      EXPORTING
        !ev_excel                TYPE xstring
        !ev_error_text           TYPE string .
    CLASS-METHODS convert_json_to_excel
      IMPORTING
        !iv_data_json            TYPE string
        !it_ddic_object          TYPE dd_x031l_table
        !it_field                TYPE za2xh_t_fieldcatalog
        !iv_sheet_title          TYPE clike OPTIONAL
        !iv_image_xstring        TYPE xstring OPTIONAL
        !iv_add_fixedvalue_sheet TYPE flag DEFAULT abap_true
        !iv_auto_column_width    TYPE flag DEFAULT abap_true
        !iv_default_descr        TYPE c DEFAULT 'L'
      EXPORTING
        !ev_excel                TYPE xstring
        !ev_error_text           TYPE string .
    CLASS-METHODS convert_excel_to_abap
      IMPORTING
        !iv_excel      TYPE xstring
        !it_field      TYPE za2xh_t_fieldcatalog OPTIONAL
        !iv_begin_row  TYPE int4 DEFAULT 2
        !iv_sheet_no   TYPE int1 DEFAULT 1
      EXPORTING
        !et_data       TYPE STANDARD TABLE
        !ev_error_text TYPE string
        !et_error_log  TYPE za2xh_t_error_log .
    CLASS-METHODS get_xstring_from_smw0
      IMPORTING
        !iv_smw0          TYPE wwwdata-objid
      RETURNING
        VALUE(rv_xstring) TYPE xstring .
  PROTECTED SECTION.

    CLASS-METHODS start_download
      IMPORTING
        !iv_excel    TYPE xstring
        !iv_filename TYPE clike OPTIONAL .
    CLASS-METHODS start_upload
      EXPORTING
        !ev_excel TYPE xstring .
    CLASS-METHODS do_drm_encode
      CHANGING
        !cv_excel TYPE xstring .
    CLASS-METHODS do_drm_decode
      CHANGING
        !cv_excel TYPE xstring .
    CLASS-METHODS add_fixedvalue_sheet
      IMPORTING
        !it_data             TYPE STANDARD TABLE
        !it_field            TYPE za2xh_t_fieldcatalog
        !it_field_catalog    TYPE zexcel_t_fieldcatalog
        !io_excel            TYPE REF TO zcl_excel
        !iv_worksheet_index  TYPE i DEFAULT 1
        !iv_header_row_index TYPE i DEFAULT 1
      RAISING
        zcx_excel .
    CLASS-METHODS add_image
      IMPORTING
        !iv_image_xstring   TYPE xstring
        !iv_col             TYPE i
        !io_excel           TYPE REF TO zcl_excel
        !iv_worksheet_index TYPE i DEFAULT 1
      RAISING
        zcx_excel .
    CLASS-METHODS get_ddic_object
      IMPORTING
        !i_data               TYPE data
      RETURNING
        VALUE(rt_ddic_object) TYPE dd_x031l_table .
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


      " 1. Get from it_field-fxied_values
      READ TABLE it_field INTO ls_field INDEX lv_index_col.
      IF sy-subrc EQ 0.
        lt_ddl = ls_field-fixed_values.
      ENDIF.

      " 2. Get from ddic domain fixed value
      IF lt_ddl IS INITIAL.
        READ TABLE lt_comp_view INTO ls_comp_view WITH KEY name = ls_field_catalog-fieldname BINARY SEARCH.
        lt_ddl = zcl_abap2xlsx_helper=>get_ddic_fixed_values( ls_comp_view-type ).
      ENDIF.

      CHECK: lt_ddl IS NOT INITIAL.
      lv_lines_ddl = lines( lt_ddl ) + 1.

      " Create fv-sheet
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

      " Add validation
      lo_data_validation = lo_worksheet->add_new_data_validation( ).
      lo_data_validation->type = zcl_excel_data_validation=>c_type_list.
      lo_data_validation->allowblank = abap_true.
      lo_data_validation->formula1 = lv_sheet_title_fv && '!$A$2:$A$' && lv_lines_ddl.
      lo_data_validation->cell_column = zcl_excel_common=>convert_column2alpha( lv_index_col ).
      lo_data_validation->cell_column_to = lo_data_validation->cell_column.
      lo_data_validation->cell_row = iv_header_row_index + 1.
      lo_data_validation->cell_row_to = lv_lines_data.

      IF iv_header_row_index IS NOT INITIAL AND lo_worksheet_fv->zif_excel_sheet_properties~hidden IS INITIAL.
        " Link to fv-sheet @ header
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
          lv_currency        TYPE tcurc-waers,
          lv_amount_external TYPE bapicurr-bapicurr,
          lv_local_ts        TYPE timestamp,
          lv_alpha_out       TYPE string,
          lv_column_count    TYPE i,
          lv_index_col       TYPE i,
          lv_index           TYPE i.
    FIELD-SYMBOLS: <ls_data>     TYPE data,
                   <lv_data>     TYPE data,
                   <lv_data_ref> TYPE data.

    CLEAR: ev_excel, ev_error_text.


    TRY.

        " Creates active sheet
        IF iv_sheet_title IS NOT INITIAL.
          lv_sheet_title = iv_sheet_title.
        ELSE.
          lv_sheet_title = 'Export'.
        ENDIF.
        CREATE OBJECT lo_excel.
        lo_worksheet = lo_excel->get_active_worksheet( ).
        lo_worksheet->set_title( ip_title = lv_sheet_title ).

        " Table settings
        ls_table_settings-table_style       = zcl_excel_table=>builtinstyle_medium2.
        ls_table_settings-show_row_stripes  = abap_true.
        ls_table_settings-nofilters         = abap_false.

        " Field catalog
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
        lv_column_count = lines( lt_field_catalog ).


**********************************************************************
        lo_worksheet->bind_table(
          EXPORTING
            ip_table          = it_data
            it_field_catalog  = lt_field_catalog
            is_table_settings = ls_table_settings
            iv_default_descr  = iv_default_descr
        ).
**********************************************************************

        IF it_ddic_object IS NOT INITIAL.
          lt_ddic_object = it_ddic_object.
        ELSE.
          lt_ddic_object = get_ddic_object( it_data ).
        ENDIF.

        " Apply conversion exit.
        LOOP AT lt_ddic_object INTO ls_ddic_object WHERE convexit IS NOT INITIAL.
          READ TABLE lt_field_catalog TRANSPORTING NO FIELDS WITH KEY fieldname = ls_ddic_object-fieldname.
          CHECK: sy-subrc EQ 0.
          lv_index_col = sy-tabix.

          IF ls_ddic_object-convexit EQ 'TSTLC' OR
             ls_ddic_object-convexit EQ 'TSTPS'.
            LOOP AT it_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_ddic_object-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              CHECK: sy-subrc EQ 0.
              IF <lv_data> IS NOT INITIAL.
                IF ls_ddic_object-convexit EQ 'TSTLC'.
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
          ELSEIF ls_ddic_object-convexit EQ 'ALPHA'.
            LOOP AT it_data ASSIGNING <ls_data>.
              lv_index = sy-tabix + 1.
              ASSIGN COMPONENT ls_ddic_object-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
              CHECK: sy-subrc EQ 0.
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

        " Apply currency
        LOOP AT lt_ddic_object INTO ls_ddic_object WHERE reffield IS NOT INITIAL.
          READ TABLE lt_field_catalog TRANSPORTING NO FIELDS WITH KEY fieldname = ls_ddic_object-fieldname.
          CHECK: sy-subrc EQ 0.
          lv_index_col = sy-tabix.

          READ TABLE lt_ddic_object INTO ls_ddic_object_ref WITH KEY fieldname = ls_ddic_object-reffield dtyp = 'CUKY'.
          CHECK: sy-subrc EQ 0.
          READ TABLE lt_field_catalog TRANSPORTING NO FIELDS WITH KEY fieldname = ls_ddic_object_ref-fieldname.
          CHECK: sy-subrc EQ 0.

          LOOP AT it_data ASSIGNING <ls_data>.
            lv_index = sy-tabix + 1.
            ASSIGN COMPONENT ls_ddic_object-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
            CHECK: sy-subrc EQ 0.
            ASSIGN COMPONENT ls_ddic_object_ref-fieldname OF STRUCTURE <ls_data> TO <lv_data_ref>.
            CHECK: sy-subrc EQ 0.
            IF <lv_data> IS NOT INITIAL AND
               <lv_data_ref> IS NOT INITIAL AND
               <lv_data_ref> NE 'USD' AND
               <lv_data_ref> NE 'EUR'.
              lv_currency = <lv_data_ref>.
              CALL FUNCTION 'BAPI_CURRENCY_CONV_TO_EXTERNAL'
                EXPORTING
                  currency        = lv_currency
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
          DO lv_column_count TIMES.
            lv_index_col = sy-index.
            lo_worksheet->set_column_width(
              EXPORTING
                ip_column         = lv_index_col
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
              iv_col             = lv_column_count + 1
              io_excel           = lo_excel
          ).
        ENDIF.

        " Freeze column headers when scrolling
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
    TYPES: BEGIN OF ts_field_conv,
             fieldname TYPE x031l-fieldname,
             convexit  TYPE x031l-convexit,
           END OF ts_field_conv,
           BEGIN OF ts_style_conv,
             cell_style TYPE zexcel_s_cell_data-cell_style,
             abap_type  TYPE abap_typekind,
           END OF ts_style_conv.
    DATA: lo_excel           TYPE REF TO zcl_excel,
          lo_reader          TYPE REF TO zif_excel_reader,
          lo_worksheet       TYPE REF TO zcl_excel_worksheet,
          lv_excel           TYPE xstring,
          ls_error_log       TYPE za2xh_s_error_log,
          lt_field           TYPE za2xh_t_fieldcatalog,
          lt_field_conv      TYPE TABLE OF ts_field_conv,
          lt_comp            TYPE abap_component_tab,
          ls_comp            TYPE abap_componentdescr,
          lo_tab_type        TYPE REF TO cl_abap_tabledescr,
          lr_data            TYPE REF TO data,
          lt_ddic_object     TYPE dd_x031l_table,
          lt_style_conv      TYPE TABLE OF ts_style_conv,
          ls_style_conv      TYPE ts_style_conv,
          ls_stylemapping    TYPE zexcel_s_stylemapping,
          lv_format_code     TYPE zexcel_number_format,
          lv_float           TYPE f,
          lt_map_excel_row   TYPE TABLE OF i,
          lv_currency        TYPE tcurc-waers,
          lv_amount_external TYPE bapicurr-bapicurr,
          lo_cx_root         TYPE REF TO cx_root,
          lo_zcx_excel       TYPE REF TO zcx_excel,
          lv_index           TYPE i,
          lv_index_col       TYPE i.
    FIELD-SYMBOLS: <lt_data>            TYPE STANDARD TABLE,
                   <ls_data>            TYPE data,
                   <lv_data>            TYPE data,
                   <lt_data2>           TYPE STANDARD TABLE,
                   <ls_data2>           TYPE data,
                   <lv_data2>           TYPE data,
                   <lv_data_ref>        TYPE data,
                   <ls_field>           TYPE za2xh_s_fieldcatalog,
                   <ls_field_conv>      TYPE ts_field_conv,
                   <ls_ddic_object>     TYPE x031l,
                   <ls_ddic_object_ref> TYPE x031l,
                   <ls_sheet_content>   TYPE zexcel_s_cell_data.

    CLEAR: ev_error_text, et_data[], et_error_log[].

    CHECK: iv_excel IS NOT INITIAL.
    lv_excel = iv_excel.
    " DRM Decode
    do_drm_decode(
      CHANGING
        cv_excel = lv_excel
    ).


    TRY.
        CREATE OBJECT lo_reader TYPE zcl_excel_reader_2007.
        lo_excel = lo_reader->load( lv_excel ).
        lo_worksheet = lo_excel->get_worksheet_by_index( iv_sheet_no ).


        " Field catalog
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

        " Create internal table with string columns
        ls_comp-type = cl_abap_elemdescr=>get_string( ).
        LOOP AT lt_field ASSIGNING <ls_field>.
          ls_comp-name = <ls_field>-fieldname.
          APPEND ls_comp TO lt_comp.
        ENDLOOP.
        lo_tab_type = cl_abap_tabledescr=>create( cl_abap_structdescr=>create( lt_comp ) ).
        CREATE DATA lr_data TYPE HANDLE lo_tab_type.
        ASSIGN lr_data->* TO <lt_data>.

        " Collect field conversion rules
        IF et_data IS SUPPLIED.
          lt_ddic_object = get_ddic_object( et_data ).
          SORT lt_ddic_object BY fieldname.
        ENDIF.
        MOVE-CORRESPONDING lt_field TO lt_field_conv.
        LOOP AT lt_field_conv ASSIGNING <ls_field_conv>.
          READ TABLE lt_ddic_object ASSIGNING <ls_ddic_object> WITH KEY fieldname = <ls_field_conv>-fieldname BINARY SEARCH.
          CHECK: sy-subrc EQ 0.
          CASE <ls_ddic_object>-exid.
            WHEN cl_abap_typedescr=>typekind_int
              OR cl_abap_typedescr=>typekind_int1
              OR cl_abap_typedescr=>typekind_int8
              OR cl_abap_typedescr=>typekind_int2
              OR cl_abap_typedescr=>typekind_packed
              OR cl_abap_typedescr=>typekind_decfloat
              OR cl_abap_typedescr=>typekind_decfloat16
              OR cl_abap_typedescr=>typekind_decfloat34
              OR cl_abap_typedescr=>typekind_float.
              " Numbers
              <ls_field_conv>-convexit = cl_abap_typedescr=>typekind_float.
            WHEN OTHERS.
              <ls_field_conv>-convexit = <ls_ddic_object>-convexit.
          ENDCASE.
        ENDLOOP.

        " Date & Time in excel style
        LOOP AT lo_worksheet->sheet_content ASSIGNING <ls_sheet_content> WHERE cell_style IS NOT INITIAL AND data_type IS INITIAL.
          ls_style_conv-cell_style = <ls_sheet_content>-cell_style.
          APPEND ls_style_conv TO lt_style_conv.
        ENDLOOP.
        IF lt_style_conv IS NOT INITIAL.
          SORT lt_style_conv BY cell_style.
          DELETE ADJACENT DUPLICATES FROM lt_style_conv COMPARING cell_style.

          LOOP AT lt_style_conv INTO ls_style_conv.

            ls_stylemapping = lo_excel->get_style_to_guid( ls_style_conv-cell_style ).
            lv_format_code = ls_stylemapping-complete_style-number_format-format_code.
            " https://support.microsoft.com/en-us/office/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68
            IF lv_format_code CS ';'.
              lv_format_code = lv_format_code(sy-fdpos).
            ENDIF.
            CHECK: lv_format_code NA '#?'.

            " Remove color pattern
            REPLACE ALL OCCURRENCES OF REGEX '\[\L[^]]*\]' IN lv_format_code WITH ''.

            IF lv_format_code CA 'yd' OR lv_format_code EQ zcl_excel_style_number_format=>c_format_date_std.
              " DATE = yyyymmdd
              ls_style_conv-abap_type = cl_abap_typedescr=>typekind_date.
            ELSEIF lv_format_code CA 'hs'.
              " TIME = hhmmss
              ls_style_conv-abap_type = cl_abap_typedescr=>typekind_time.
            ELSE.
              DELETE lt_style_conv.
              CONTINUE.
            ENDIF.

            MODIFY lt_style_conv FROM ls_style_conv TRANSPORTING abap_type.

          ENDLOOP.
        ENDIF.


**********************************************************************
* Start of convert content
**********************************************************************
        READ TABLE lo_worksheet->sheet_content TRANSPORTING NO FIELDS WITH KEY cell_row = iv_begin_row.
        IF sy-subrc EQ 0.
          lv_index = sy-tabix.
        ENDIF.

        LOOP AT lo_worksheet->sheet_content ASSIGNING <ls_sheet_content> FROM lv_index.
          AT NEW cell_row.
            " New line
            APPEND INITIAL LINE TO <lt_data> ASSIGNING <ls_data>.
            lv_index = sy-tabix.
          ENDAT.

          IF <ls_sheet_content>-cell_value IS NOT INITIAL.
            ASSIGN COMPONENT <ls_sheet_content>-cell_column OF STRUCTURE <ls_data> TO <lv_data>.
            IF sy-subrc EQ 0.
              " value
              <lv_data> = <ls_sheet_content>-cell_value.

              " field conversion
              READ TABLE lt_field_conv ASSIGNING <ls_field_conv> INDEX <ls_sheet_content>-cell_column.
              IF sy-subrc EQ 0 AND <ls_field_conv>-convexit IS NOT INITIAL.
                CASE <ls_field_conv>-convexit.
                  WHEN cl_abap_typedescr=>typekind_float.
                    lv_float = zcl_excel_common=>excel_string_to_number( <ls_sheet_content>-cell_value ).
                    <lv_data> = |{ lv_float NUMBER = RAW }|.
                  WHEN 'ALPHA'.
                    CALL FUNCTION 'CONVERSION_EXIT_ALPHA_OUTPUT'
                      EXPORTING
                        input  = <ls_sheet_content>-cell_value
                      IMPORTING
                        output = <lv_data>.
                ENDCASE.
              ENDIF.

              " style conversion
              IF <ls_sheet_content>-cell_style IS NOT INITIAL.
                READ TABLE lt_style_conv INTO ls_style_conv WITH KEY cell_style = <ls_sheet_content>-cell_style BINARY SEARCH.
                IF sy-subrc EQ 0.
                  CASE ls_style_conv-abap_type.
                    WHEN cl_abap_typedescr=>typekind_date.
                      <lv_data> = zcl_excel_common=>excel_string_to_date( <ls_sheet_content>-cell_value ).
                    WHEN cl_abap_typedescr=>typekind_time.
                      <lv_data> = zcl_excel_common=>excel_string_to_time( <ls_sheet_content>-cell_value ).
                  ENDCASE.
                ENDIF.
              ENDIF.
              CONDENSE <lv_data>.
            ENDIF.
          ENDIF.

          AT END OF cell_row.
            " Delete empty line
            IF <ls_data> IS INITIAL.
              DELETE <lt_data> INDEX lv_index.
            ELSE.
              APPEND <ls_sheet_content>-cell_row TO lt_map_excel_row.
            ENDIF.
          ENDAT.
        ENDLOOP.
**********************************************************************
* End of convert content
**********************************************************************


        CHECK: <lt_data> IS NOT INITIAL.
        MOVE-CORRESPONDING <lt_data> TO et_data.


        " Find errors
        IF et_error_log IS SUPPLIED.
          CREATE DATA lr_data LIKE <lt_data>.
          ASSIGN lr_data->* TO <lt_data2>.
          MOVE-CORRESPONDING et_data TO <lt_data2>.
          LOOP AT lt_field_conv ASSIGNING <ls_field_conv>
           WHERE convexit = cl_abap_typedescr=>typekind_float.
            LOOP AT <lt_data2> ASSIGNING <ls_data2>.
              ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data2> TO <lv_data2>.
              CHECK: sy-subrc EQ 0 AND <lv_data> IS NOT INITIAL.
              lv_float = <lv_data2>.
              <lv_data2> = |{ lv_float NUMBER = RAW }|.
            ENDLOOP.
          ENDLOOP.

          IF <lt_data> NE <lt_data2>.
            LOOP AT <lt_data> ASSIGNING <ls_data>.
              lv_index = sy-tabix.
              READ TABLE <lt_data2> ASSIGNING <ls_data2> INDEX lv_index.
              IF <ls_data> NE <ls_data2>.
                LOOP AT lt_field_conv ASSIGNING <ls_field_conv>.
                  lv_index_col = sy-tabix.
                  ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
                  CHECK: sy-subrc EQ 0.
                  ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data2> TO <lv_data2>.
                  CHECK: sy-subrc EQ 0.
                  IF <lv_data> NE <lv_data2>.
                    CLEAR: ls_error_log.
                    ls_error_log-row = lv_index.
                    ls_error_log-fieldname = <ls_field_conv>-fieldname.
                    ls_error_log-abap_value = <lv_data2>.
                    ls_error_log-excel_value = <lv_data>.
                    READ TABLE lt_map_excel_row INTO lv_index INDEX lv_index.
                    READ TABLE lo_worksheet->sheet_content ASSIGNING <ls_sheet_content> WITH TABLE KEY cell_row = lv_index cell_column = lv_index_col.
                    ls_error_log-excel_coords = <ls_sheet_content>-cell_coords.
                    APPEND ls_error_log TO et_error_log.
                  ENDIF.
                ENDLOOP.
              ENDIF.
            ENDLOOP.
          ENDIF.
        ENDIF.


        " Apply conversion exit.
        LOOP AT lt_field_conv ASSIGNING <ls_field_conv>
         WHERE convexit = 'ALPHA'
            OR convexit = 'TSTLC'.
          LOOP AT et_data ASSIGNING <ls_data>.
            ASSIGN COMPONENT <ls_field_conv>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
            CHECK: sy-subrc EQ 0 AND <lv_data> IS NOT INITIAL.
            CASE <ls_field_conv>-convexit.
              WHEN 'ALPHA'.
                CALL FUNCTION 'CONVERSION_EXIT_ALPHA_INPUT'
                  EXPORTING
                    input  = <lv_data>
                  IMPORTING
                    output = <lv_data>.
              WHEN 'TSTLC'.
                PERFORM convert_to_utc_time IN PROGRAM saplsdc_cnv USING <lv_data> CHANGING <lv_data>.
            ENDCASE.
          ENDLOOP.
        ENDLOOP.


        " Apply currency
        LOOP AT lt_ddic_object ASSIGNING <ls_ddic_object> WHERE reffield IS NOT INITIAL.
          READ TABLE lt_ddic_object ASSIGNING <ls_ddic_object_ref> WITH KEY fieldname = <ls_ddic_object>-reffield dtyp = 'CUKY'.
          CHECK: sy-subrc EQ 0.
          LOOP AT et_data ASSIGNING <ls_data>.
            ASSIGN COMPONENT <ls_ddic_object>-fieldname OF STRUCTURE <ls_data> TO <lv_data>.
            CHECK: sy-subrc EQ 0.
            IF <lv_data> IS NOT INITIAL.
              ASSIGN COMPONENT <ls_ddic_object_ref>-fieldname OF STRUCTURE <ls_data> TO <lv_data_ref>.
              CHECK: sy-subrc EQ 0.
              IF <lv_data_ref> IS NOT INITIAL AND
                 <lv_data_ref> NE 'USD' AND
                 <lv_data_ref> NE 'EUR'.
                lv_currency = <lv_data_ref>.
                lv_amount_external = <lv_data>.
                CALL FUNCTION 'BAPI_CURRENCY_CONV_TO_INTERNAL'
                  EXPORTING
                    currency             = lv_currency
                    amount_external      = lv_amount_external
                    max_number_of_digits = 23
                  IMPORTING
                    amount_internal      = <lv_data>.
              ENDIF.
            ENDIF.
          ENDLOOP.
        ENDLOOP.

      CATCH zcx_excel INTO lo_zcx_excel.    " Exceptions for ABAP2XLSX
        ev_error_text = lo_zcx_excel->error.
      CATCH cx_root INTO lo_cx_root.
        ev_error_text = lo_cx_root->get_text( ).
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
        et_error_log  = et_error_log
    ).

    IF ev_error_text IS NOT INITIAL.
      zcl_abap2xlsx_helper=>message( iv_error_text = ev_error_text ).
    ENDIF.

  ENDMETHOD.


  METHOD get_ddic_object.
    DATA: lo_type_desc   TYPE REF TO cl_abap_typedescr,
          lo_table_desc  TYPE REF TO cl_abap_tabledescr,
          lo_struct_desc TYPE REF TO cl_abap_structdescr,
          lt_comp_view   TYPE abap_component_view_tab,
          ls_comp_view   TYPE abap_simple_componentdescr,
          lt_ddic_object TYPE dd_x031l_table,
          ls_ddic_object TYPE x031l.

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
          not_found    = 1
          no_ddic_type = 2
          OTHERS       = 3
      ).
      IF rt_ddic_object IS INITIAL.
        lt_comp_view = lo_struct_desc->get_included_view( ).
        LOOP AT lt_comp_view INTO ls_comp_view.
          ls_comp_view-type->get_ddic_object(
            RECEIVING
              p_object     = lt_ddic_object
            EXCEPTIONS
              not_found    = 1
              no_ddic_type = 2
              OTHERS       = 3
          ).
          IF lt_ddic_object IS NOT INITIAL.
            READ TABLE lt_ddic_object INTO ls_ddic_object INDEX 1.
            ls_ddic_object-fieldname = ls_comp_view-name.
            APPEND ls_ddic_object TO rt_ddic_object.
          ENDIF.
        ENDLOOP.
      ENDIF.
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
    CONSTANTS: lc_xlsx_mime TYPE string VALUE 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'.
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
      cl_wd_runtime_services=>attach_file_to_response(
        EXPORTING
          i_filename      = lv_filename_string
          i_content       = iv_excel
          i_mime_type     = lc_xlsx_mime
      ).
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
      IF sy-subrc NE 0.
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
        IF sy-subrc NE 0.
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
      IF sy-subrc NE 0.
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
      IF sy-subrc NE 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.

      CALL FUNCTION 'SCMS_BINARY_TO_XSTRING'
        EXPORTING
          input_length = lv_filelength
        IMPORTING
          buffer       = ev_excel
        TABLES
          binary_tab   = lt_temptable
        EXCEPTIONS
          failed       = 1
          OTHERS       = 2.
      IF sy-subrc NE 0.
        MESSAGE ID sy-msgid TYPE sy-msgty NUMBER sy-msgno
          WITH sy-msgv1 sy-msgv2 sy-msgv3 sy-msgv4.
      ENDIF.
    ENDIF.
  ENDMETHOD.
ENDCLASS.
