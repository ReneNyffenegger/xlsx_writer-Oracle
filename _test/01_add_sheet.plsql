declare

  workbook xlsx_writer.book_r;
  sheet    integer;

  xlsx     blob;

begin

  workbook := xlsx_writer.start_book;
  sheet    := xlsx_writer.add_sheet  (workbook, 'Name of the sheet');
  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '01_add_sheet.xlsx', xlsx);

end;
/

@after_test_item 01_add_sheet
