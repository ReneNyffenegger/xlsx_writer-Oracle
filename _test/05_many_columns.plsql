declare

  workbook xlsx_writer.book_r;
  sheet        integer;

  xlsx          blob;

  font_courier  integer;
  cs_courier    integer;

  font_bold     integer;
  cs_bold       integer;

  font_italic   integer;
  cs_italic     integer;

  font_underl   integer;
  cs_underl     integer;

begin

  workbook     := xlsx_writer.start_book;
  sheet        := xlsx_writer.add_sheet     (workbook, 'Name of the sheet');

  for col in 1 .. 100 loop
     xlsx_writer.add_cell(workbook, sheet, 1, col, value_  => col); 
  end loop;

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '05_many_columns.xlsx', xlsx);

end;
/

@after_test_item 05_many_columns

