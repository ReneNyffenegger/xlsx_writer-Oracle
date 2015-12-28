declare

  workbook xlsx_writer.book_r;
  sheet_1  integer;
  sheet_2  integer;
  sheet_3  integer;

  xlsx     blob;

  c_limit   constant integer := 50;
  c_x_split constant integer := 2;
  c_y_split constant integer := 3;

begin

  workbook := xlsx_writer.start_book;
  sheet_1  := xlsx_writer.add_sheet  (workbook, 'Name of sheet one');

  -- { First sheet

  for i in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet_1,   c_y_split, i+c_x_split, value_ => i);
  end loop;

  for i in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet_1, i+c_y_split,   c_x_split, value_ => i);
  end loop;

  for x in 1 .. c_limit loop
  for y in 1 .. c_limit loop
      xlsx_writer.add_cell(workbook, sheet_1, c_y_split + x, c_x_split + y, value_ => x*y);
  end loop;
  end loop;

  xlsx_writer.freeze_sheet(workbook, sheet_1, c_x_split, c_y_split);

  -- }
  -- { Second sheet
  sheet_2  := xlsx_writer.add_sheet  (workbook, 'Name of sheet two');

  xlsx_writer.add_cell(workbook, sheet_2, 1, 1, text => 'Foo');
  xlsx_writer.add_cell(workbook, sheet_2, 1, 2, text => 'Bar');
  xlsx_writer.add_cell(workbook, sheet_2, 1, 3, text => 'Baz');

  for r in 1 .. 100 loop
    xlsx_writer.add_cell(workbook, sheet_2, r+1, 1, value_ => r);
    xlsx_writer.add_cell(workbook, sheet_2, r+1, 2, value_ => r*1001);
    xlsx_writer.add_cell(workbook, sheet_2, r+1, 3, value_ => r*10010001);
  end loop;

  xlsx_writer.col_width(workbook, sheet_2, 3, 16);
  xlsx_writer.freeze_sheet(workbook, sheet_2, split_y => 1);

  -- }
  -- { Third sheet
  sheet_3  := xlsx_writer.add_sheet  (workbook, 'Name of sheet three');

  xlsx_writer.add_cell(workbook, sheet_3, 1, 1, text => 'Foo');
  xlsx_writer.add_cell(workbook, sheet_3, 2, 1, text => 'Bar');
  xlsx_writer.add_cell(workbook, sheet_3, 3, 1, text => 'Baz');

  for r in 1 .. 100 loop
    xlsx_writer.add_cell(workbook, sheet_3, 1, r+1, value_ => r);
    xlsx_writer.add_cell(workbook, sheet_3, 2, r+1, value_ => r*1001);
    xlsx_writer.add_cell(workbook, sheet_3, 3, r+1, value_ => r*10010001);
    xlsx_writer.col_width(workbook, sheet_3, r+1, 16);
  end loop;

  xlsx_writer.freeze_sheet(workbook, sheet_3, split_x => 1);

  -- }

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '07_freeze_sheet.xlsx', xlsx);

end;
/

@after_test_item 07_freeze_sheet
