create table tq84_xlsx_writer_test (
  col_1 varchar2(20),
  col_2 number,
  col_3 date
);

insert into tq84_xlsx_writer_test values ('foo',   42     , date '2016-01-11' + 13/17);
insert into tq84_xlsx_writer_test values ('bar', null     , date '2016-01-12'        );
insert into tq84_xlsx_writer_test values ('baz', -108.2094, null                     );
insert into tq84_xlsx_writer_test values ( null,    0     , date '1970-01-01'        );

declare

  xlsx blob;

begin

  xlsx := xlsx_writer.sql_to_xlsx('select * from tq84_xlsx_writer_test order by col_1');
  blob_wrapper.to_file('XLSX_WRITER_TEST_DIR', 's2e_01.xlsx', xlsx);

end;
/

drop table tq84_xlsx_writer_test purge;

@after_test_item s2e_01
