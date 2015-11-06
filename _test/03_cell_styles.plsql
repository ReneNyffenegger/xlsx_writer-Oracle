declare

  workbook xlsx_writer.book_r;
  sheet                   integer;

  xlsx                    blob;

  font_courier            integer;
  cs_courier              integer;

  font_bold               integer;
  cs_bold                 integer;

  font_italic             integer;
  cs_italic               integer;

  font_underl             integer;
  cs_underl               integer;

  num_fmt_dd_mm_yyyy      integer;
  cs_dd_mm_yyyy           integer;

begin

  workbook     := xlsx_writer.start_book;
  sheet        := xlsx_writer.add_sheet     (workbook, 'Name of the sheet');
  font_courier := xlsx_writer.add_font      (workbook, 'Courier New', 12);
  font_bold    := xlsx_writer.add_font      (workbook, 'Arial'      , 12, b => true);
  font_italic  := xlsx_writer.add_font      (workbook, 'Georgia'    , 12, i => true);
  font_underl  := xlsx_writer.add_font      (workbook, 'Verdana'    , 12, u => true);

  xlsx_writer.col_width(workbook, sheet, 1, 1, 28);
  xlsx_writer.col_width(workbook, sheet, 2, 2, 14);

  cs_courier   := xlsx_writer.add_cell_style(workbook, font_id => font_courier);
  cs_bold      := xlsx_writer.add_cell_style(workbook, font_id => font_bold);
  cs_italic    := xlsx_writer.add_cell_style(workbook, font_id => font_italic);
  cs_underl    := xlsx_writer.add_cell_style(workbook, font_id => font_underl);

  num_fmt_dd_mm_yyyy := xlsx_writer.add_num_fmt_dd_mm_yyyy(workbook);
  cs_dd_mm_yyyy      := xlsx_writer.add_cell_style(workbook, num_fmt_id => num_fmt_dd_mm_yyyy);


  xlsx_writer.add_cell(workbook, sheet, 1, 1, style_id => cs_courier, text => 'Courier'   ); 
  xlsx_writer.add_cell(workbook, sheet, 2, 1, style_id => cs_bold   , text => 'Bold'      ); 
  xlsx_writer.add_cell(workbook, sheet, 3, 1, style_id => cs_italic , text => 'Italic'    ); 
  xlsx_writer.add_cell(workbook, sheet, 4, 1, style_id => cs_underl , text => 'Underlined'); 

  xlsx_writer.add_cell(workbook, sheet, 5, 1, text => 'Date in dd.mm.yyyy format'); 
  xlsx_writer.add_cell(workbook, sheet, 5, 2, style_id => cs_dd_mm_yyyy, date_ => date '2015-06-05'); 

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '03_cell_styles.xlsx', xlsx);

end;
/

@after_test_item 03_cell_styles
