# xlsx writer for Oracle

Creating Excel workbooks (xlsx) with PL/SQL

This project was forked from [xlsx_writer-Oracle](https://github.com/ReneNyffenegger/xlsx_writer-Oracle).

Basically I improve cell fill function, font color, horizontal alignment and improve performance in generation large files.  
To improve performance in my version I limited use of dbms_lob.append using it instead a buffering system on varchar2,   
in this way performance improves on large files up to 95%.   

## Cell fill  

I modify oracle function add_fill, passing two parameters type_ (solid) and color (rgb format) :  

*fill1 := xlsx_writer.add_fill(workbook, type_ => 'solid', color => '999999');*  

## Font Color  

I modify parameter color (rgb format) in function add_font :  

*font_bold1    := xlsx_writer.add_font      (workbook, 'Arial'      , 24, color => '000000', b => true );*  

## Horizontal alignment  

I introduce a new parameter horizontal_alignment in function add_cell_style :  

*cs_fill1 := xlsx_writer.add_cell_style(workbook, horizontal_alignment => 'center', fill_id => fill1, font_id => font_bold1);*  

# Dependencies

The package needs `zipper` which can be found [here](https://github.com/ReneNyffenegger/oracle_scriptlets/tree/master/zipper).  

The [tests](https://github.com/ReneNyffenegger/xlsx_writer-Oracle/tree/master/_test) require the
[`blob_writer`](https://github.com/ReneNyffenegger/blob_wrapper-Oracle) package.

# Links

[About Office Open XML](https://github.com/ReneNyffenegger/about-Office-Open-XML).

[Examples on renenyffenegger homepage](http://renenyffenegger.ch/Oracle/Libraries/xlsx-writer.html).

Example for the procedure [sql_to_xlsx](http://renenyffenegger.blogspot.ch/2016/01/oracle-turning-select-statement-into.html) which turns an SQL statement into an excel sheet.

# Licence

xlsx writer is licenced under the **GNU General Public License v3.0**.
