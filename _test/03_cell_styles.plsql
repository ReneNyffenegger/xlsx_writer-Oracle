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


 "mm-dd-yy"               integer;
 "0"                      integer;
 "0.00"                   integer;
 "#.##0"                  integer;
 "#.##0.00"               integer;
 "0%"                     integer;
 "0.00%"                  integer;
 "h:mm:ss"                integer;

begin

  workbook     := xlsx_writer.start_book;
  sheet        := xlsx_writer.add_sheet     (workbook, 'Name of the sheet');
  font_courier := xlsx_writer.add_font      (workbook, 'Courier New', 12);
  font_bold    := xlsx_writer.add_font      (workbook, 'Arial'      , 12, b => true);
  font_italic  := xlsx_writer.add_font      (workbook, 'Georgia'    , 12, i => true);
  font_underl  := xlsx_writer.add_font      (workbook, 'Verdana'    , 12, u => true);

  xlsx_writer.col_width(workbook, sheet, 1, 1, 45);
  xlsx_writer.col_width(workbook, sheet, 2, 2, 14);

  cs_courier   := xlsx_writer.add_cell_style(workbook, font_id => font_courier);
  cs_bold      := xlsx_writer.add_cell_style(workbook, font_id => font_bold);
  cs_italic    := xlsx_writer.add_cell_style(workbook, font_id => font_italic);
  cs_underl    := xlsx_writer.add_cell_style(workbook, font_id => font_underl);

 "mm-dd-yy"          := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."mm-dd-yy");
 "0"                 := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."0"       );
 "0.00"              := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."0.00"    );
 "#.##0"             := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."#.##0"   );
 "#.##0.00"          := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."#.##0.00");
 "0%"                := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."0%"      );
 "0.00%"             := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."0.00%"   );
 "h:mm:ss"           := xlsx_writer.add_cell_style(workbook, num_fmt_id => xlsx_writer."h:mm:ss" );


  xlsx_writer.add_cell(workbook, sheet,  1, 1, style_id => cs_courier, text => 'Courier'   ); 
  xlsx_writer.add_cell(workbook, sheet,  2, 1, style_id => cs_bold   , text => 'Bold'      ); 
  xlsx_writer.add_cell(workbook, sheet,  3, 1, style_id => cs_italic , text => 'Italic'    ); 
  xlsx_writer.add_cell(workbook, sheet,  4, 1, style_id => cs_underl , text => 'Underlined'); 

  xlsx_writer.add_cell(workbook, sheet,  5, 1, text => 'Date in mm-dd-yy'); 
  xlsx_writer.add_cell(workbook, sheet,  5, 2, style_id => "mm-dd-yy", date_ => date '2015-06-05'); 

  xlsx_writer.add_cell(workbook, sheet,  6, 1, text => 'Number 12345.678 in "0"'); 
  xlsx_writer.add_cell(workbook, sheet,  6, 2, style_id => "0", value_ => 12345.678); 

  xlsx_writer.add_cell(workbook, sheet,  7, 1, text => 'Number 12345.678 in "0.00"'); 
  xlsx_writer.add_cell(workbook, sheet,  7, 2, style_id => "0.00", value_ => 12345.678); 

  xlsx_writer.add_cell(workbook, sheet,  8, 1, text => 'Number 12345.678 in "#.##0"'); 
  xlsx_writer.add_cell(workbook, sheet,  8, 2, style_id => "#.##0", value_ => 12345.678); 

  xlsx_writer.add_cell(workbook, sheet,  9, 1, text => 'Number 12345.678 in "#.##0.00"'); 
  xlsx_writer.add_cell(workbook, sheet,  9, 2, style_id => "#.##0.00", value_ => 12345.678); 

  xlsx_writer.add_cell(workbook, sheet, 10, 1, text => 'Number 0.1783 in "0%"'); 
  xlsx_writer.add_cell(workbook, sheet, 10, 2, style_id => "0%", value_ => 0.1783); 

  xlsx_writer.add_cell(workbook, sheet, 11, 1, text => 'Number 0.1783 in "0.00%"'); 
  xlsx_writer.add_cell(workbook, sheet, 11, 2, style_id => "0.00%", value_ => 0.1783); 

  xlsx_writer.add_cell(workbook, sheet, 12, 1, text => 'Number 21/24+13/24/60+48/24/60/60 in "h:mm:ss"'); 
  xlsx_writer.add_cell(workbook, sheet, 12, 2, style_id => "h:mm:ss", value_ => 21/24+13/24/60+48/24/60/60); 

  xlsx     := xlsx_writer.create_xlsx(workbook);

  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', '03_cell_styles.xlsx', xlsx);

end;
/

@after_test_item 03_cell_styles
