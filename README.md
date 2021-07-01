# ABAP2XLSX_HELPER
abap2xlsx helper

[**Korean**](#korean)
&nbsp;·&nbsp;
[**English**](#english)

## Korean
아직 개발중입니다. 일부 동작합니다.
- ZCL_ABAP2XLSX_HELPER=>EXCEL_DOWNLOAD
<BR>인터널 테이블 내용을 엑셀 파일로 다운로드 합니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>EXCEL_EMAIL
<BR>인터널 테이블 내용을 엑셀 파일로 첨부하여 이메일을 발송 합니다. 변환작업을 백그라운드로 처리하므로 실행은 바로 끝납니다. 용량에 따라 이메일이 늦게 도착 할 수 있습니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>EXCEL_UPLOAD
<BR>엑셀 파일을 업로드하여 인터널 테이블에 넣습니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>GET_FIELDCATALOG
- ZCL_ABAP2XLSX_HELPER=>CONVERT_ABAP_TO_EXCEL
<BR>인터널 테이블 내용을 엑셀 파일(XSTRING)로 변환 합니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>CONVERT_JSON_TO_EXCEL
<BR>JSON 표현의 데이터를 읽어 엑셀 파일(XSTRING)로 변환 합니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>CONVERT_EXCEL_TO_ABAP
<BR>엑셀 파일(XSTRING)을 읽어 인터널 테이블에 넣습니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>GET_XSTRING_FROM_SMW0
<BR>SMW0에 업로드한 이미지 파일을 읽어서 XSTRING으로 변환 합니다. 엑셀 다운로드시 IV_IMAGE_XSTRING 파라미터로 전달하면 이미지를 추가 할 수 있습니다.<BR>
- ZCL_ABAP2XLSX_HELPER=>FPM_UPLOAD_POPUP
- ZCL_ABAP2XLSX_HELPER=>GET_DDIC_FIXED_VALUES

## English
this is not stable. you can use these methods.
- ZCL_ABAP2XLSX_HELPER=>EXCEL_DOWNLOAD
<BR>convert abap internal table to excel file and start to download.<BR>
- ZCL_ABAP2XLSX_HELPER=>EXCEL_EMAIL
<BR>convert abap internal table to excel file and send by email. execution is finished quickly. converting is processed in background. in case of large convertion, email delays can occur<BR>
- ZCL_ABAP2XLSX_HELPER=>EXCEL_UPLOAD
<BR>start to upload file and convert excel file to abap internal table.<BR>
- ZCL_ABAP2XLSX_HELPER=>GET_FIELDCATALOG
- ZCL_ABAP2XLSX_HELPER=>CONVERT_ABAP_TO_EXCEL
<BR>convert abap internal table to excel XSTRING<BR>
- ZCL_ABAP2XLSX_HELPER=>CONVERT_JSON_TO_EXCEL
<BR>convert JSON data(it must come from abap internal table) to excel XSTRING<BR>
- ZCL_ABAP2XLSX_HELPER=>CONVERT_EXCEL_TO_ABAP
<BR>convert excel XSTRING to abap internal table.<BR>
- ZCL_ABAP2XLSX_HELPER=>GET_XSTRING_FROM_SMW0
<BR>convert image file from SMW0 to XSTRING. if add image to excel, use IV_IMAGE_XSTRING parameter.<BR>
- ZCL_ABAP2XLSX_HELPER=>FPM_UPLOAD_POPUP
- ZCL_ABAP2XLSX_HELPER=>GET_DDIC_FIXED_VALUES
