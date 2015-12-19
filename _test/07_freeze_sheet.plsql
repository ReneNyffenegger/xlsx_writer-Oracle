declare

  workbook xlsx_writer.book_r;
  sheet    integer;

  xlsx     blob;

  c_limit   constant integer := 50;
  c_x_split constant integer := 2;
  c_y_split constant integer := 3;

begin

  workbook := xlsx_writer.start_book;
  sheet    := xlsx_writer.add_sheet  (workbook, 'Name of the sheet');

  for i in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet,   c_y_split, i+c_x_split, value_ => i);
  end loop;

  for i in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet, i+c_y_split,   c_x_split, value_ => i);
  end loop;

  for x in 1 .. c_limit loop
  for y in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet, c_y_split + x, c_x_split + y, value_ => x*y);
  end loop;
  end loop;

  xlsx_writer.freeze_sheet(workbook, sheet, c_x_split, c_y_split);

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '07_freeze_sheet.xlsx', xlsx);

end;
/

@after_test_item 07_freeze_sheet
