create or replace package xlsx_writer as -- {{{
-- vi: foldmarker={{{,}}}

  "0"                          constant integer :=  1;
  "0.00"                       constant integer :=  2;
  "#.##0"                      constant integer :=  3;
  "#.##0.00"                   constant integer :=  4;

  "0%"                         constant integer :=  9;
  "0.00%"                      constant integer := 10;
  "0.00E+00"                   constant integer := 11;
  "# ?/?"                      constant integer := 12;
  "# ??/??"                    constant integer := 13;
  "mm-dd-yy"                   constant integer := 14; -- Note, for a reason, this is displayed as «dd.mm.yyyy», at least with my «locale» settings.
  "d-mmm-yy"                   constant integer := 15;
  "d-mmm"                      constant integer := 16;
  "mmm-yy"                     constant integer := 17;
  "h:mm AM/PM"                 constant integer := 18;
  "h:mm:ss AM/PM"              constant integer := 19;
  "h:mm"                       constant integer := 20;
  "h:mm:ss"                    constant integer := 21;
  "m/d/yy h:mm"                constant integer := 22;

  "#,##0 ;(#,##0)"             constant integer := 37;
  "#,##0 ;[Red](#,##0)"        constant integer := 38;
  "#,##0.00;(#,##0.00)"        constant integer := 39;
  "#,##0.00;[Red](#,##0.00)"   constant integer := 40;

  "mm:ss"                      constant integer := 45;
  "[h]:mm:ss"                  constant integer := 46;
  "mmss.0"                     constant integer := 47;
  "##0.0E+0"                   constant integer := 48;
  "@"                          constant integer := 49;



  -- {{{ Types

  -- {{{ Related to styles

  -- {{{ Borders

  type border_r             is record(raw_        varchar2(1000));
  type border_t             is table of border_r;

  -- }}}

  -- {{{ Fills

  type fill_r               is record(patternType varchar2(20),
                                      rgb         varchar2(20)
                                     );
  type fill_t               is table of fill_r;


  -- }}}

  -- {{{ Font
  type font_r               is record(name        varchar2(100),
                                      size_       number       ,
                                      color       varchar2(100), -- «theme="1"»  or  «rgb="FFFF0000"»
                                      b           boolean      ,
                                      i           boolean      ,
                                      u           boolean
                                  --  family_val  number         -- <family val="2" /> 
                                    );

  type font_t               is table of font_r;
  -- }}}

  -- {{{ Num Formats

  type num_fmt_r            is record(raw_        varchar2(1000));
  type num_fmt_t            is table of num_fmt_r;


  -- }}}

  -- {{{ Cell Styles

  type cell_style_r         is record(font_id              integer,
                                      fill_id              integer,
                                      border_id            integer,
                                      num_fmt_id           integer,
                                      vertical_alignment   varchar2(10),
                                      horizontal_alignment varchar2(10),
                                      wrap_text            boolean
                                  --  raw_within varchar2(200)
                                    );

  type cell_style_t         is table of cell_style_r;

  -- }}}

  -- }}}

  -- {{{ shared Strings

  type shared_string_r      is record(val        varchar2(4000));
  type shared_string_t      is table of shared_string_r;

  -- }}}

  -- {{{ Sheet, consisting of rows, consisting of cells / columns have their width

  -- {{{ Column Widths

  type col_width_r          is record(start_col        integer,
                                      end_col          integer,
                                      width            number
                               --     style            number
                               );
  type col_width_t          is table of col_width_r;

  -- }}}

  -- {{{ Cell

  type cell_r               is record(style_id         integer,
                                      shared_string_id integer,
                                      value_           number,
                                      formula          varchar2(4000));
               
  type cell_t               is table of cell_r index by pls_integer;

  -- }}}

  -- {{{ Row

  type row_r                is record(r          integer,
                                      height     number,
                                      cells      cell_t);

  type row_t                is table of row_r index by pls_integer;

  -- }}}

  -- {{{ Worksheet relations

  type sheet_rel_r          is record(raw_       varchar2(4000));
  type sheet_rel_t          is table of sheet_rel_r;

  -- }}}

  -- {{{ vml-Drawings

  type checkbox_r           is record(text               varchar2(4000),
                                      checked            boolean,
                                      col_left           integer,
                                      row_top            integer);

  type checkbox_t           is table of checkbox_r;

  type vml_drawing_r        is record(checkboxes          checkbox_t);

  type vml_drawing_t        is table of vml_drawing_r;

  -- }}}

  -- {{{ Sheet
  type sheet_r              is record(col_widths     col_width_t,
                                      name_          varchar2(100),
                                      rows_          row_t,
                                      split_x        integer,
                                      split_y        integer,
                                      sheet_rels     sheet_rel_t,
                                      vml_drawings   vml_drawing_t -- Can there be multiple drawings per sheet? (ASSMPT_01)?
                               );

  type sheet_t              is table of sheet_r;

  -- }}}

  -- }}}

  -- {{{ Medias

  type media_r              is record(b                blob,
                                      name_            varchar2(100));

  type media_t              is table of media_r;

  -- }}}

  -- {{{ Calc Chain

  type calc_chain_elem_r    is record(cell_reference   varchar2(10),
                                      sheet            integer);

  type calc_chain_elem_t    is table of calc_chain_elem_r;

  -- }}}

  -- {{{ Drawings

  type drawing_r            is record(raw_             varchar2(30000));

  type drawing_t            is table of drawing_r;

  -- }}}


  -- {{{ The book!

  type book_r               is record(sheets                   sheet_t,
                                      cell_styles              cell_style_t,
                                      borders                  border_t,
                                      fonts                    font_t,
                                      fills                    fill_t,
                                      num_fmts                 num_fmt_t,
                                      shared_strings           shared_string_t,
                                      medias                   media_t,
                                      calc_chain_elems         calc_chain_elem_t,
                                      drawings                 drawing_t,
                                      content_type_vmlDrawing  boolean
                                      );

  -- }}}

  -- }}}

  function  start_book                                       return book_r;
  
  function  add_sheet         (xlsx         in out book_r,
                               name_        in     varchar2) return integer;

  function add_cell_style     (xlsx                 in out book_r,
                               font_id              integer  := 0,
                               fill_id              integer  := 0,
                               border_id            integer  := 0,
                               num_fmt_id           integer  := 0,
                               vertical_alignment   varchar2 := null,
                               horizontal_alignment varchar2 := null,
                               wrap_text            boolean  := null
                           --  raw_within          varchar2 := null
                            ) return integer;

  function add_border         (xlsx         in out book_r,
                               raw_                varchar2) return integer;

  function add_num_fmt        (xlsx         in out book_r,
                               raw_                varchar2,
                               return_id           integer) return integer;

--function add_num_fmt_0          (xlsx    in out book_r)  return integer;
--function add_num_fmt_dd_mm_yyyy (xlsx    in out book_r)  return integer; -- Date in «dd.mm.yyyy» format


  function add_font           (xlsx         in out book_r,
                               name                varchar2,
                               size_               number,
                               color               varchar2   := null,
                               b                   boolean    := false,
                               i                   boolean    := false,
                               u                   boolean    := false) return integer;

  function add_fill           (xlsx         in out book_r,
                               type_        varchar2 := 'solid',
                               color        varchar2 := '000000' 
                               ) return integer;
          
  procedure col_width         (xlsx         in out book_r,
                               sheet              integer,
                               col                integer,
                               width              number
                               );

  procedure col_width         (xlsx         in out book_r,
                               sheet              integer,
                               start_col          integer,
                               end_col            integer,
                               width              number
                       --      style              number
                               );

  procedure add_row           (xlsx        in out book_r,
                               sheet              integer,
                               r                  integer,
                               height             number := null);

  procedure freeze_sheet      (xlsx        in out book_r,
                               sheet       in     integer,
                               split_x     in     integer := null,
                               split_y     in     integer := null);
                             

  procedure add_cell          (xlsx        in out book_r,
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               style_id           integer  :=    0,
                               text               varchar2 := null,
                               value_             number   := null,
                               formula            varchar2 := null,
                               height             number  := null);

  procedure add_cell          (xlsx        in out book_r,
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               date_              date,
                               style_id           integer  :=    0);

  procedure add_sheet_rel     (xlsx        in out book_r,
                               sheet              integer,
                               raw_               varchar2);

  procedure add_media         (xlsx        in out book_r,
                               b                  blob,
                               name_              varchar2);        

  procedure add_drawing       (xlsx        in out book_r,
                               raw_               varchar2);

  procedure add_checkbox      (xlsx        in out book_r,
                               sheet              integer,
                               col_left           integer,
                               row_top            integer,
                               text               varchar2 := null,
                               checked            boolean  := false);
                          

  function create_xlsx        (xlsx        in out book_r) return blob;

  function col_to_letter(c integer) return varchar2;

  function sql_to_xlsx(sql_stmt varchar2) return blob;

end xlsx_writer; -- }}}
/
show errors
