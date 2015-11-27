
declare

  workbook xlsx_writer.book_r;
  sheet    integer;

  xlsx     blob;

begin

  workbook := xlsx_writer.start_book;
  sheet    := xlsx_writer.add_sheet  (workbook, 'Name of the sheet');

  xlsx_writer.add_row(workbook, sheet, 2);

  xlsx_writer.col_width(workbook, sheet, 3, 4, 20);

  xlsx_writer.add_checkbox(workbook, sheet, 3, 2, 'Check me');
  xlsx_writer.add_checkbox(workbook, sheet, 4, 2, 'Uncheck me', checked=>true);

  xlsx     := xlsx_writer.create_xlsx(workbook);


  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '06_controls.xlsx', xlsx);

end;
/

@after_test_item 06_controls
