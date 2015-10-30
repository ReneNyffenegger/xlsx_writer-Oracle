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


  xlsx_writer.add_cell(workbook, sheet, 1, 1, value_  => 1); 
  xlsx_writer.add_cell(workbook, sheet, 2, 1, value_  => 1); 
  xlsx_writer.add_cell(workbook, sheet, 3, 1, formula => 'SUM(A1:A2)', value_ =>  2); 
  xlsx_writer.add_cell(workbook, sheet, 4, 1, formula => 'SUM(A2:A3)', value_ =>  3); 
  xlsx_writer.add_cell(workbook, sheet, 5, 1, formula => 'SUM(A3:A4)', value_ =>  5); 
  xlsx_writer.add_cell(workbook, sheet, 6, 1, formula => 'SUM(A4:A5)', value_ =>  8); 
  xlsx_writer.add_cell(workbook, sheet, 7, 1, formula => 'SUM(A5:A6)', value_ => 13); 
  xlsx_writer.add_cell(workbook, sheet, 8, 1, formula => 'SUM(A6:A7)', value_ => 21); 

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '04_formulas.xlsx', xlsx);

end;
/

@after_test_item 04_formulas

