FUNCTION za2xh_email.
*"----------------------------------------------------------------------
*"*"Local Interface:
*"  IMPORTING
*"     REFERENCE(IT_DATA) TYPE  TABLE
*"     REFERENCE(IT_FIELD) TYPE  ZA2XH_T_FIELDCATALOG OPTIONAL
*"     REFERENCE(IV_SUBJECT) TYPE  CLIKE OPTIONAL
*"     REFERENCE(IV_SENDER) TYPE  CLIKE OPTIONAL
*"     REFERENCE(IT_RECEIVER) TYPE  STRINGTAB
*"     REFERENCE(IV_FILENAME) TYPE  CLIKE OPTIONAL
*"     REFERENCE(IV_SHEET_TITLE) TYPE  CLIKE OPTIONAL
*"     REFERENCE(IV_IMAGE_XSTRING) TYPE  XSTRING OPTIONAL
*"     REFERENCE(IV_ADD_FIXEDVALUE_SHEET) TYPE  FLAG DEFAULT ABAP_TRUE
*"     REFERENCE(IV_AUTO_COLUMN_WIDTH) TYPE  FLAG DEFAULT ABAP_TRUE
*"     REFERENCE(IV_DEFAULT_DESCR) TYPE  C DEFAULT 'L'
*"----------------------------------------------------------------------
  DATA: lt_field           TYPE za2xh_t_fieldcatalog,
        lv_xstring         TYPE xstring,
        lv_filename_string TYPE string,
        lv_parti           TYPE i VALUE 50000,
        lv_count           TYPE i,
        lv_from            TYPE i,
        lv_to              TYPE i,
        lr_data            TYPE REF TO data,
        ls_comp            TYPE abap_componentdescr,
        lt_comp            TYPE abap_component_tab,
        lo_tab_type        TYPE REF TO cl_abap_tabledescr,
        lo_struc_type      TYPE REF TO cl_abap_structdescr,
        lv_data_json       TYPE string,
        lt_comp_view       TYPE abap_component_view_tab,
        ls_comp_view       TYPE abap_simple_componentdescr,
        lt_ddic_object     TYPE dd_x031l_table,
        lv_index           TYPE i.
*  CLEAR ev_error_text.
  FIELD-SYMBOLS: <lt_table>       TYPE table,
                 <lt_table_parti> TYPE table,
                 <ls_field>       TYPE za2xh_s_fieldcatalog.

  IF it_field IS NOT INITIAL.
    lt_field = it_field.
  ELSE.
    zcl_abap2xlsx_helper=>get_fieldcatalog(
      EXPORTING
        it_data          = it_data
*        iv_default_descr = 'L'
      IMPORTING
        et_field         = lt_field
    ).
  ENDIF.

  lo_tab_type ?= cl_abap_tabledescr=>describe_by_data( it_data ).
  lo_struc_type ?= lo_tab_type->get_table_line_type( ).
  lo_struc_type->get_ddic_object(
    RECEIVING
      p_object     = lt_ddic_object
    EXCEPTIONS
      not_found    = 1
      no_ddic_type = 2
      others       = 3
  ).
  lt_comp_view = lo_struc_type->get_included_view( ).
  SORT lt_comp_view BY name.
  LOOP AT lt_field ASSIGNING <ls_field>.
    READ TABLE lt_comp_view INTO ls_comp_view WITH KEY name = <ls_field>-fieldname BINARY SEARCH.
    IF sy-subrc <> 0.
      DELETE lt_field.
      CONTINUE.
    ENDIF.

    IF <ls_field>-fixed_values IS INITIAL.
      <ls_field>-fixed_values = zcl_abap2xlsx_helper=>get_ddic_fixed_values( ls_comp_view-type ).
    ENDIF.

    ls_comp-name = <ls_field>-fieldname.
    ls_comp-type = cl_abap_elemdescr=>get_string( ).
    APPEND ls_comp TO lt_comp.
  ENDLOOP.



  CHECK: lt_comp IS NOT INITIAL.
  lo_tab_type = cl_abap_tabledescr=>create( cl_abap_structdescr=>create( lt_comp ) ).
  CREATE DATA lr_data TYPE HANDLE lo_tab_type.
  ASSIGN lr_data->* TO <lt_table>.
  MOVE-CORRESPONDING it_data[] TO <lt_table>.

  CREATE DATA lr_data LIKE <lt_table>.
  ASSIGN lr_data->* TO <lt_table_parti>.



  lv_count = lines( <lt_table> ).
  lv_from = 1.
  WHILE lv_from <= lv_count.
    lv_index = lv_index + 1.
    lv_to = lv_from + lv_parti - 1.
    CLEAR: <lt_table_parti>, lv_filename_string, lv_data_json.
    INSERT LINES OF <lt_table> FROM lv_from TO lv_to INTO TABLE <lt_table_parti>.
    /ui2/cl_json=>serialize(
      EXPORTING
        data             = <lt_table_parti>
      RECEIVING
        r_json           = lv_data_json
    ).


    IF iv_filename IS NOT INITIAL.
      lv_filename_string = iv_filename.
    ELSE.
      lv_filename_string = zcl_abap2xlsx_helper=>default_excel_filename( ).
    ENDIF.
    IF lv_index EQ 1 AND lv_to >= lv_count.
      " 파일 1개.
    ELSE.
      REPLACE '.xlsx' IN lv_filename_string WITH '' IGNORING CASE.
      lv_filename_string = |{ lv_filename_string }_part{ lv_index }.xlsx|.
    ENDIF.


    CALL FUNCTION 'ZA2XH_EMAIL_RFC'
      IN BACKGROUND TASK AS SEPARATE UNIT
      EXPORTING
        iv_data_json            = lv_data_json
        it_ddic_object          = lt_ddic_object
        it_field                = lt_field
        iv_subject              = CONV string( iv_subject )
        iv_sender               = CONV string( iv_sender )
        it_receiver             = it_receiver
        iv_filename             = lv_filename_string
        iv_sheet_title          = CONV string( iv_sheet_title )
        iv_image_xstring        = iv_image_xstring
        iv_add_fixedvalue_sheet = iv_add_fixedvalue_sheet
        iv_auto_column_width    = iv_auto_column_width
        iv_default_descr        = iv_default_descr.

    lv_from = lv_from + lv_parti.
  ENDWHILE.

  COMMIT WORK.

ENDFUNCTION.
