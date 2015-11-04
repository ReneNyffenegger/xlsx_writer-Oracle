set define &
define test_out_dir=c:\temp
define unzip_cmd=c:\about\Office-Open-XML\_helpers\office_unzip.pl

create directory xlsx_writer_test_dir as '&test_out_dir';

@01_add_sheet.plsql
@02_fill_cells.plsql
@03_cell_styles.plsql
@04_formulas.plsql
@05_many_columns.plsql

drop   directory xlsx_writer_test_dir;
