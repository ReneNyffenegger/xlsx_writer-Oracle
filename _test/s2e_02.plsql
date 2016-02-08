create table tq84_xlsx_writer_s2e_test (
  col_001   varchar2(  1),
  col_004   varchar2(  4),
  col_010   varchar2( 10),
  col_019   varchar2( 19),
  col_030   varchar2( 31),
  col_050   varchar2( 50),
  col_100   varchar2(100),
  col_002   varchar2(  2)
);


insert into tq84_xlsx_writer_s2e_test values (
  'x',
  'four',
  'ten things',
  'nineteen characters',
  'abcdef ghijkl mno pqrst uvw xyz',
  'Looking for a text consisting of fifty characters.',
  '123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789!',
  '12'
);

declare

  xlsx blob;

begin

  xlsx := xlsx_writer.sql_to_xlsx('select * from tq84_xlsx_writer_s2e_test');
  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', 's2e_02.xlsx', xlsx);

end;
/

drop table tq84_xlsx_writer_s2e_test purge;

@after_test_item s2e_02
