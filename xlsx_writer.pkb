set define off
create or replace package body xlsx_writer as -- {{{
-- vi: foldmarker={{{,}}}

  procedure ap (b in out blob, v in varchar2) is -- {{{
  begin
    if  v is not null then 
        dbms_lob.append(b, utl_raw.cast_to_raw(v));
    end if;
  end ap; -- }}}

  procedure aps (b in out blob, v1 in out varchar2, v2 in varchar2) is -- {{{
  begin
    if length(v1) + length(v2) > 4000 then
       ap(b,v1);
       v1 := null;
    end if;   
    v1 := v1 || v2;
  end aps; -- }}}

  function start_xml_blob return blob is -- {{{
    ret blob;
  begin
    dbms_lob.createTemporary(ret, true);
    
    ap(ret,q'{<?xml version="1.0" encoding="utf-8"?>
}');

    return ret;

  end start_xml_blob; -- }}}

  procedure add_attr(b          in out blob, -- {{{
                     attr_name         varchar2,
                     attr_value        varchar2) is
  begin
    ap(b, ' ' || attr_name || '="' || attr_value || '"');
  end add_attr; -- }}}

  procedure add_attrs(b          in out blob,
                      v          in out varchar2, -- {{{
                      attr_name         varchar2,
                      attr_value        varchar2) is
  begin
    if length(v) + 1 + length(attr_name) + 1 + 1 + length(attr_value) + 1 > 4000 then
       ap(b,v);
       v := null;
    end if;
    v := v || ' ' || attr_name || '="' || attr_value || '"';
  end add_attrs; -- }}}


  procedure warning(text varchar2) is -- {{{
  begin
    dbms_output.put_line('! warnining: ' || text);
  end warning; -- }}}

  function start_book return book_r is -- {{{
    ret book_r;
  begin

    ret.sheets                   := new sheet_t           ();

    ret.cell_styles              := new cell_style_t      ();
    ret.borders                  := new border_t          ();
    ret.fonts                    := new font_t            ();
    ret.fills                    := new fill_t            ();
    ret.num_fmts                 := new num_fmt_t         ();
    ret.shared_strings           := new shared_string_t   ();
    ret.medias                   := new media_t           ();
    ret.calc_chain_elems         := new calc_chain_elem_t ();
    ret.drawings                 := new drawing_t         ();
    ret.content_type_vmlDrawing  := false;

    return ret;

  end start_book; -- }}}

  function  add_sheet(xlsx in out book_r, -- {{{
                      name_       varchar2) return integer is

    ret sheet_r;
  begin

 -- Sheetname must not contain any of : [ ]
    ret.name_      := translate(name_, ':[]', '   ');

    ret.col_widths := new col_width_t();
    ret.sheet_rels := new sheet_rel_t();

    xlsx.sheets.extend;
    xlsx.sheets(xlsx.sheets.count) := ret;

    return xlsx.sheets.count;

  end add_sheet; -- }}}

  procedure freeze_sheet      (xlsx        in out book_r, -- {{{
                               sheet       in     integer,
                               split_x     in     integer := null,
                               split_y     in     integer := null) is
  begin

    xlsx.sheets(sheet).split_x := split_x;
    xlsx.sheets(sheet).split_y := split_y;

  end freeze_sheet; -- }}}

  procedure add_sheet_rel     (xlsx        in out book_r, -- {{{
                               sheet              integer,
                               raw_               varchar2)
  is
  begin

    xlsx.sheets(sheet).sheet_rels.extend;
    xlsx.sheets(sheet).sheet_rels(xlsx.sheets(sheet).sheet_rels.count). raw_ := raw_;

  end add_sheet_rel; -- }}}

  procedure add_media         (xlsx        in out book_r, -- {{{
                               b                  blob,
                               name_              varchar2) is
  begin
  
    xlsx.medias.extend;
    xlsx.medias(xlsx.medias.count).b     := b;
    xlsx.medias(xlsx.medias.count).name_ := name_;


  end add_media; -- }}}

  procedure add_drawing       (xlsx        in out book_r, -- {{{
                               raw_               varchar2) is
  begin

    xlsx.drawings.extend;
    xlsx.drawings(xlsx.drawings.count).raw_ := raw_;
  end add_drawing; -- }}}

  procedure add_checkbox      (xlsx        in out book_r, -- {{{
                               sheet              integer,
                               col_left           integer,
                               row_top            integer,
                               text               varchar2 := null,
                               checked            boolean  := false) is
  begin

--  TODO hier weiter machen: die checkbox in die xlsx.vml_drawings einfügen.

    if xlsx.sheets(sheet).vml_drawings is null then
       xlsx.sheets(sheet).vml_drawings := new vml_drawing_t();

       xlsx.sheets(sheet).vml_drawings.extend;

    -- Assumption ASSMPT_01: at most ONE vml drawing per sheet.
    -- index set to 1.
       xlsx.sheets(sheet).vml_drawings(1).checkboxes := new checkbox_t();
    end if;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes.extend;
    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).col_left := col_left;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).row_top  := row_top;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).text     := text;

    xlsx.sheets(sheet).vml_drawings(1).checkboxes(
    xlsx.sheets(sheet).vml_drawings(1).checkboxes.count).checked  := checked;

    xlsx.content_type_vmlDrawing := true;

  end add_checkbox; -- }}}

  -- {{{ Rows and Columns

  function col_to_letter(c integer) return varchar2 is -- {{{
  begin

    if c < 27 then
       return substr('ABCDEFGHIJKLMNOPQRSTUVWXYZ', c, 1);
    end if;

    return col_to_letter(trunc((c-1)/26)) || col_to_letter(mod((c-1), 26)+1);
    
  end col_to_letter; -- }}}

  procedure col_width         (xlsx    in out book_r,-- {{{
                               sheet          integer,
                               col            integer,
                               width          number
                               ) is
  begin

    col_width(xlsx, sheet, col, col, width);


  end col_width; -- }}}

  procedure col_width         (xlsx    in out book_r,-- {{{
                               sheet          integer,
                               start_col      integer,
                               end_col        integer,
                               width          number
                          --   style          number
                               ) is
    r col_width_r;
  begin

    r.start_col := start_col;
    r.end_col   := end_col;
    r.width     := width;
--  r.style     := style;


    xlsx.sheets(sheet).col_widths.extend;
    xlsx.sheets(sheet).col_widths(xlsx.sheets(sheet).col_widths.count) := r;

  end col_width; -- }}}

  function does_row_exist(xlsx  in out book_r, -- {{{
                          sheet        integer,
                          r            integer) return boolean is
  begin

     if xlsx.sheets(sheet).rows_.exists(r) then
        return true;
     end if;

     return false;

  end does_row_exist; -- }}}

  procedure add_row(xlsx     in out book_r, -- {{{
                    sheet           integer,
                    r               integer,
                    height          number := null) is
  begin

      if does_row_exist(xlsx, sheet, r) then
         raise_application_error(-20800, 'row ' || r || ' already exists');
      end if;

      xlsx.sheets(sheet).rows_(r).height := height;

  end add_row; -- }}}

  procedure add_cell          (xlsx    in out book_r, -- {{{
                               sheet         integer,
                               r             integer,
                               c             integer,
                               style_id      integer  :=    0,
                               text          varchar2 := null,
                               value_        number   := null,
                               formula       varchar2 := null,
                               height        number := null) is
  begin

    if not does_row_exist(xlsx, sheet, r) then
       add_row(xlsx, sheet, r, height);
    end if;

    if xlsx.sheets(sheet).rows_(r).cells.exists(c) then
       warning('Cell ' || c || ' in row ' || r || ' already exists.');
    end if;

    if style_id is null then
       raise_application_error(-20800, 'style id is null for cell ' || r || '/' || c);
    end if;

    xlsx.sheets(sheet).rows_(r).cells(c).style_id := style_id;
    xlsx.sheets(sheet).rows_(r).cells(c).value_   := value_;
    xlsx.sheets(sheet).rows_(r).cells(c).formula  := formula;

    if formula is not null then -- {{{

       xlsx.calc_chain_elems.extend;
       xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).cell_reference := col_to_letter(c) || r;
       xlsx.calc_chain_elems(xlsx.calc_chain_elems.count).sheet          := sheet;

    end if; -- }}}

    if text is not null then -- {{{
       xlsx.shared_strings.extend;
       xlsx.shared_strings(xlsx.shared_strings.count).val := replace(
                                                             replace(
                                                             replace(text, '&', '&amp;'),
                                                                           '>', '&gt;' ),
                                                                           '<', '&lt;' );
 
       xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id := xlsx.shared_strings.count-1;
    end if; -- }}}

  end add_cell; -- }}}

  procedure add_cell          (xlsx        in out book_r, -- {{{
                               sheet              integer,
                               r                  integer,
                               c                  integer,
                               date_              date,
                               style_id           integer  :=    0) is
  begin

    add_cell(xlsx         => xlsx,
             sheet        => sheet,
             r            => r,
             c            => c,
             style_id     => style_id,
             value_       => date_ - date '1899-12-30');

  end add_cell; -- }}}

  -- }}}

  -- {{{ Related to styles
  --
  function add_font         (xlsx     in out book_r, -- {{{
                             name            varchar2,
                             size_           number,
                             color           varchar2 := null,
                             b               boolean  := false,
                             i               boolean  := false,
                             u               boolean  := false) return integer is
  begin

      xlsx.fonts.extend;
      xlsx.fonts(xlsx.fonts.count).name  := name; 
      xlsx.fonts(xlsx.fonts.count).size_ := size_; 
      xlsx.fonts(xlsx.fonts.count).color := color; 
      xlsx.fonts(xlsx.fonts.count).b     := b; 
      xlsx.fonts(xlsx.fonts.count).i     := i; 
      xlsx.fonts(xlsx.fonts.count).u     := u; 

      return xlsx.fonts.count; -- Not returning xlsx.fonts.count because of default font.

  end add_font; -- }}}

  function add_fill        (xlsx     in out book_r, -- {{{
                            type_    varchar2 := 'solid',
                            color    varchar2 := '000000') return integer is

  begin

    xlsx.fills.extend;
    xlsx.fills(xlsx.fills.count).patternType := type_;
    xlsx.fills(xlsx.fills.count).rgb := color;

    return xlsx.fills.count + 1;  -- +1 instead of -1 because there are two default fills.

  end add_fill; -- }}}

  function add_border      (xlsx     in out book_r, -- {{{
                            raw_            varchar2) return integer is
  begin

    xlsx.borders.extend;
    xlsx.borders(xlsx.borders.count).raw_ := raw_;

    return xlsx.borders.count; -- Return count instead of count-1 because the empty border is default
  end add_border; -- }}}

  function add_cell_style  (xlsx                 in out book_r, -- {{{
                            font_id              integer  := 0,
                            fill_id              integer  := 0,
                            border_id            integer  := 0,
                            num_fmt_id           integer  := 0,
                            vertical_alignment   varchar2 := null,
                            horizontal_alignment varchar2 := null,
                            wrap_text            boolean  := null
                         -- raw_within          varchar2 := null
                          ) return integer is

    rec cell_style_r;
  begin

    rec.font_id              := font_id;
    rec.fill_id              := fill_id;
    rec.border_id            := border_id;
    rec.num_fmt_id           := num_fmt_id;
--  rec.raw_within           := raw_within;
    rec.vertical_alignment   := vertical_alignment;
    rec.horizontal_alignment := horizontal_alignment;
    rec.wrap_text            := wrap_text;

    xlsx.cell_styles.extend;
    xlsx.cell_styles(xlsx.cell_styles.count) := rec;

    return xlsx.cell_styles.count; -- Not returning count -1 because of «default style»

  end add_cell_style; -- }}}

  function add_num_fmt     (xlsx     in out book_r, -- {{{
                            raw_            varchar2,
                            return_id       integer) return integer is

  begin

    xlsx.num_fmts.extend;
    xlsx.num_fmts(xlsx.num_fmts.count).raw_ := raw_;

--  return xlsx.num_fmts.count - 1;
    return return_id;

  end add_num_fmt; -- }}}

  /*
  function add_num_fmt_0          (xlsx    in out book_r)  return integer is -- {{{ "0"
  begin
    return 1;
  end add_num_fmt_0; -- }}}
  function add_num_fmt_dd_mm_yyyy (xlsx    in out book_r)  return integer is -- {{{ Date in «dd.mm.yyyy» format
  begin

    return 14;

  end add_num_fmt_dd_mm_yyyy; -- }}}
  */
  
  -- }}}

  -- {{{ Blob generators

  function docProps_app return blob is -- {{{
    ret blob;
    rets varchar2(4000);
  begin

    ret := start_xml_blob;
    rets := null;

    aps(ret, rets, '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">');
    aps(ret, rets, '</Properties>');
    ap(ret, rets);

    return ret;

  end docProps_app; -- }}}

  function xl_worksheets_sheet(xlsx in out book_r, -- {{{
                              sheet        integer) return blob is

    ret blob;
    rets varchar2(4000);
  begin

    ret := start_xml_blob;
    rets := null;

    aps(ret,rets, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">');

    if xlsx.sheets(sheet).split_x is not null or -- {{{
       xlsx.sheets(sheet).split_y is not null then

       aps(ret, rets, '<sheetViews>');        
       aps(ret, rets, '<sheetView workbookViewId="0">');        

       aps(ret, rets, '<pane');

       if xlsx.sheets(sheet).split_x is not null then
          add_attrs(ret, rets, 'xSplit', xlsx.sheets(sheet).split_x);
       end if; 

       if xlsx.sheets(sheet).split_y is not null then
          add_attrs(ret, rets, 'ySplit', xlsx.sheets(sheet).split_y);
       end if; 

       add_attrs(ret, rets, 'topLeftCell', col_to_letter(nvl(xlsx.sheets(sheet).split_x, 1) + 1) || (nvl(xlsx.sheets(sheet).split_y, 1) + 1));
       aps(ret, rets, ' state="frozen"/>');

       if xlsx.sheets(sheet).split_y is not null then
          aps(ret, rets, '<selection pane="bottomLeft" activeCell="B6" sqref="B6" />');
       else
          aps(ret, rets, '<selection pane="topRight" activeCell="B6" sqref="B6" />');
       end if;

       aps(ret, rets, '</sheetView>');        
       aps(ret, rets, '</sheetViews>');        

    end if; -- }}}

    if xlsx.sheets(sheet).col_widths.count > 0 then -- {{{
      aps(ret, rets, '<cols>');

      for i in 1 .. xlsx.sheets(sheet).col_widths.count loop -- {{{

        aps(ret, rets, '<col');

        add_attrs(ret, rets, 'min'        , xlsx.sheets(sheet).col_widths(i).start_col);
        add_attrs(ret, rets, 'max'        , xlsx.sheets(sheet).col_widths(i).end_col  );
        add_attrs(ret, rets, 'width'      , xlsx.sheets(sheet).col_widths(i).width    );
--      add_attr(ret, 'style'      , xlsx.sheets(sheet).col_widths(i).style    );
        add_attrs(ret, rets, 'customWidth', 1);

        aps (ret, rets, '/>');

      end loop; -- }}}

      aps(ret, rets, '</cols>');
    end if; -- }}}

    aps(ret, rets, '<sheetData>'); -- {{{

    declare
      r pls_integer;
      c pls_integer;
    begin
      r := xlsx.sheets(sheet).rows_.first;
      while r is not null loop -- {{{
 
        aps(ret, rets, '<row');
        add_attrs(ret, rets, 'r', r);

        if xlsx.sheets(sheet).rows_(r).height is not null then
           add_attrs(ret, rets, 'ht', xlsx.sheets(sheet).rows_(r).height);
           add_attrs(ret, rets, 'customHeight', 1);
        end if;

        aps(ret, rets, '>');
 
/*      if xlsx.sheets(sheet).rows_(r).cells.count = 0 then
           raise_application_error(-20800, 'Row ' || r || ' does not contain any cells');
        end if;*/


        c := xlsx.sheets(sheet).rows_(r).cells.first;
        while c is not null loop -- {{{

          aps(ret, rets, '<c');

          add_attrs(ret, rets, 'r', col_to_letter(c) || r);
          add_attrs(ret, rets, 's', xlsx.sheets(sheet).rows_(r).cells(c).style_id);

          if xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id is not null then
             add_attrs(ret, rets, 't', 's'); -- Type is String
          end if;

          aps(ret, rets, '>');


          if xlsx.sheets(sheet).rows_(r).cells(c).formula is not null then
             aps(ret, rets, '<f>' || xlsx.sheets(sheet).rows_(r).cells(c).formula || '</f>');
          end if;

          if xlsx.sheets(sheet).rows_(r).cells(c).value_ is not null then
             aps(ret, rets, '<v>' || xlsx.sheets(sheet).rows_(r).cells(c).value_ || '</v>');
          end if;

          if xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id is not null then
             aps(ret, rets, '<v>' || xlsx.sheets(sheet).rows_(r).cells(c).shared_string_id || '</v>');
          end if;

          aps(ret, rets, '</c>');

          c := xlsx.sheets(sheet).rows_(r).cells.next(c);

        end loop; -- }}}
        
        r := xlsx.sheets(sheet).rows_.next(r);
 
        aps(ret, rets, '</row>');
      end loop; -- }}}
    end;

    aps(ret, rets, '</sheetData>'); -- }}}

    aps(ret, rets, '<pageMargins left="0.36000000000000004" right="0.2" top="1" bottom="1" header="0.5" footer="0.5" />');
  
    aps(ret, rets, '<pageSetup paperSize="9" scale="67" orientation="landscape" horizontalDpi="4294967292" verticalDpi="4294967292" />');
  
    for d in 1 .. xlsx.drawings.count loop 
      aps(ret, rets, '<drawing r:id="rId' || d || '" /> ');
    end loop;

    if xlsx.sheets(sheet).vml_drawings is not null then
       aps(ret, rets, '<legacyDrawing r:id="rel_vml_drawing_' || sheet || '" />');
    end if;

    aps(ret, rets, '</worksheet>');
    ap(ret, rets);

    return ret;

  end xl_worksheets_sheet; -- }}}

  function xl_styles(xlsx in book_r) return blob is -- {{{
    ret blob;
    rets varchar2(4000);
  begin
    ret := start_xml_blob;
    rets := null;

    aps(ret, rets, '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">');

    if xlsx.num_fmts.count > 0 then -- {{{
       aps(ret, rets, '<numFmts>');

       for n in 1 .. xlsx.num_fmts.count loop
           aps(ret, rets, '<numFmt ' || xlsx.num_fmts(n).raw_ || ' />');
       end loop;

       aps(ret, rets, '</numFmts>');

    end if; -- }}}

    aps(ret, rets, '<fonts>'); -- {{{

    aps(ret, rets, '<font><sz val="11"/><name val="Calibri" /></font>'); -- Default font.

    for f in 1 .. xlsx.fonts.count loop -- {{{

      aps(ret, rets, '<font><name val="' || xlsx.fonts(f).name  || '"/>' ||
                    '  <sz val="' || xlsx.fonts(f).size_ || '"/>');

      if xlsx.fonts(f).color is not null then
         aps(ret, rets, '<color rgb="' || UPPER(xlsx.fonts(f).color) || '"/>');
      end if;

      if xlsx.fonts(f).i then
         aps(ret, rets, '<i/>');
      end if;

      if xlsx.fonts(f).b then
         aps(ret, rets, '<b/>');
      end if;

      if xlsx.fonts(f).u then
         aps(ret, rets, '<u/>');
      end if;

      aps(ret, rets, '</font>');

    end loop; -- }}}

    aps(ret, rets, '</fonts>'); -- }}}

    aps(ret, rets, '<fills count="' || (xlsx.fills.count + 2)  || '">'); -- {{{

--  the first two pattern fills seem to somehow be default...
    aps(ret, rets, q'{
      <fill><patternFill patternType="none"    /></fill>
      <fill><patternFill patternType="gray125" /></fill>
    }');

    for f in 1 .. xlsx.fills.count loop -- {{{

      aps(ret, rets, '<fill><patternFill patternType="' || xlsx.fills(f).patternType || '"><fgColor rgb="' || UPPER(xlsx.fills(f).rgb) || '"/></patternFill></fill>');

    end loop; -- }}}

    aps(ret, rets, '</fills>'); -- }}}

    aps(ret, rets, '<borders>'); -- {{{

    -- Add default «empty» border
    aps(ret, rets, '<border><left/><right/><top/><bottom/><diagonal/></border>');


    for b in 1 .. xlsx.borders.count loop -- {{{
        aps(ret, rets, '<border>' || xlsx.borders(b).raw_ || '</border>');
    end loop; -- }}}

    aps(ret, rets, '</borders>'); -- }}}

    -- <cellStyleXfs>  Huh, what's this for? -- {{{
    aps(ret, rets, q'{<cellStyleXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
    </cellStyleXfs>}'); -- }}}

    aps(ret, rets, '<cellXfs>'); -- {{{ Cell Styles

 -- Default cell style?
    aps(ret, rets, '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />');

    for c in 1 .. xlsx.cell_styles.count loop -- {{{

        aps(ret, rets, '<xf');

        add_attrs(ret, rets, 'numFmtId', xlsx.cell_styles(c).num_fmt_id);
        add_attrs(ret, rets, 'fillId'  , xlsx.cell_styles(c).fill_id   );
        add_attrs(ret, rets, 'fontId'  , xlsx.cell_styles(c).font_id   );
        add_attrs(ret, rets, 'borderId', xlsx.cell_styles(c).border_id );
        aps(ret, rets, '>');

        if xlsx.cell_styles(c).vertical_alignment is not null OR
           xlsx.cell_styles(c).horizontal_alignment is not null OR 
           xlsx.cell_styles(c).wrap_text
        then
               aps(ret, rets, '<alignment ');
               if xlsx.cell_styles(c).vertical_alignment is not null
               then
                   aps(ret, rets, 'vertical="' || xlsx.cell_styles(c).vertical_alignment || '" ');
               end if;

               if xlsx.cell_styles(c).horizontal_alignment is not null
               then
                  aps(ret, rets, 'horizontal="' || xlsx.cell_styles(c).horizontal_alignment || '" ');
               end if;

               if xlsx.cell_styles(c).wrap_text is not null 
               then
                  aps(ret, rets, ' wrapText="' || case when xlsx.cell_styles(c).wrap_text then '1' else '0' end || '"');
               end if;
               aps(ret, rets, '/>');
        end if;

 --     if xlsx.cell_styles(c).raw_within is not null then
 --        ap(ret, xlsx.cell_styles(c).raw_within);
 --     end if;

        aps(ret, rets, '</xf>');

    end loop; -- }}}

  aps(ret, rets, '</cellXfs>'); -- }}}

    aps(ret, rets, q'{<cellStyles count="33">
      <cellStyle name="Normal" xfId="0" builtinId="0" />
   </cellStyles>}');

    aps(ret, rets, q'{<dxfs count="0" />}');


    aps(ret, rets, q'{</styleSheet>}');
    ap(ret, rets);

    return ret;

  end xl_styles; -- }}}

  function xl_sharedStrings(xlsx book_r) return blob is -- {{{
    ret blob;
    rets varchar2(4000);
  begin
    ret := start_xml_blob();
    rets := null;

    aps(ret, rets, '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="24" uniqueCount="24">');

    for s in 1 .. xlsx.shared_strings.count loop
      aps(ret, rets, '<si><t xml:space="preserve">' || xlsx.shared_strings(s).val || '</t></si>');
    end loop;

    aps(ret, rets, '</sst>');
    ap(ret, rets);
    return ret;
  end xl_sharedStrings; -- }}}

  function xl_workbook(xlsx in out book_r -- {{{
  ) return blob is
    ret blob := start_xml_blob;
    rets varchar2(4000);
  begin

    rets := null;
    aps(ret, rets, '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'); -- {{{

--  This line obviously necessary to prevent «has changed» message when xlsx
--  is closed and contains formulas.
    aps(ret, rets, '  <fileVersion appName="xl" lastEdited="5" lowestEdited="5" rupBuild="9303" />'); -- Older Excel version has lastEdited="4"

    aps(ret, rets, '<sheets>'); -- {{{

    for s in 1 .. xlsx.sheets.count loop -- {{{

        aps(ret, rets, '<sheet');

        add_attrs(ret, rets, 'name'   , xlsx.sheets(s).name_);
        add_attrs(ret, rets, 'sheetId', s                   );
        add_attrs(ret, rets, 'r:id'   ,'rId' || s           );

        aps(ret, rets, '/>');

     end loop; -- }}}

    aps(ret, rets, '</sheets>'); -- }}}

--  This line obviously necessary to prevent «has changed» message when xlsx
--  is closed and contains formulas.
    aps(ret, rets, '<calcPr calcId="145621" />'); -- Older Excel Version: calcId="125725"

--  ap(ret, '<calcPr calcOnSave="0" />');

    aps(ret, rets, '</workbook>'); -- }}}
    ap(ret, rets);

    return ret;
  
  end xl_workbook; -- }}}

  function rels_rels      (xlsx in out book_r) return blob is -- {{{
    ret blob;
    rets varchar2(4000);
  begin

    ret := start_xml_blob();
    rets := null;

    aps(ret, rets, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
    aps(ret, rets, '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"  />');
    aps(ret, rets, '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"   Target="docProps/core.xml" />');
    aps(ret, rets, '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"      Target="xl/workbook.xml"   />');
    aps(ret, rets, '</Relationships>');

    ap(ret, rets);
    return ret;

  end rels_rels; -- }}}

  function docProps_core  (xlsx in out book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000);
  begin
    rets := null;

    aps(ret, rets, '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">');
    aps(ret, rets, '</cp:coreProperties>');
    ap(ret, rets);

    return ret;

  end docProps_core; -- }}}

  function xl_rels_workbook           (xlsx in out book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000) := null;
    rId integer := 1;

    procedure add_relationship(type_ varchar2, target varchar2) is -- {{{
    begin
      
      aps(ret, rets,  '<Relationship');

      add_attrs(ret, rets, 'Id'    , 'rId' || rId); rId := rId + 1;
      add_attrs(ret, rets, 'Type'  ,  type_      );
      add_attrs(ret, rets, 'Target',  target     );

      aps(ret, rets, '/>');


    end add_relationship; -- }}}
  begin

    rets := null;

    aps(ret, rets, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');

      for s in 1 .. xlsx.sheets.count loop
        add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', 'worksheets/sheet' || s || '.xml');
      end loop;

      add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings', 'sharedStrings.xml');
      add_relationship('http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'       , 'styles.xml'       );

    aps(ret, rets, '</Relationships>');
    ap(ret,rets);
    return ret;

  end xl_rels_workbook; -- }}}

  function xl_drawings_rels_drawing1  (xlsx in out book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000) := null;
  begin

    aps(ret, rets, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');

    for m in 1 .. xlsx.medias.count loop
      aps(ret, rets, '<Relationship Id="rId' || m || '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' || xlsx.medias(m).name_ || '" />');
    end loop;

    aps(ret, rets, '</Relationships>');
    ap(ret, rets);

    return ret;

  end xl_drawings_rels_drawing1; -- }}}

  function xl_worksheets_rels_sheet(xlsx in out book_r, sheet integer)  return blob is -- {{{
    ret blob;
    rets varchar2(4000);
  begin

    rets := null;

    for r in 1 .. xlsx.sheets(sheet).sheet_rels.count loop

      if ret is null then
         ret := start_xml_blob;
         aps(ret, rets, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
      end if;

      aps(ret, rets, xlsx.sheets(sheet).sheet_rels(r).raw_);

    end loop;

--  if xlsx.calc_chain_elems.count > 0 then
--     if ret is null then
--        ret := start_xml_blob;
--        ap(ret, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
--     end if;
--     ap(ret, '<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml" />');
--  end if;


    if xlsx.sheets(sheet).vml_drawings is not null then
       if ret is null then
          ret := start_xml_blob;
          aps(ret, rets, '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">');
       end if;
       aps(ret, rets, '<Relationship Id="rel_vml_drawing_' || sheet ||  '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing' || sheet || '.vml" />');
    end if;

    if ret is not null then
       aps(ret, rets, '</Relationships>');
    end if;

    ap(ret,rets);

    return ret;
  end xl_worksheets_rels_sheet; -- }}}

  function content_types(xlsx in out book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000) := null;
  begin

    aps(ret, rets, '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">');


    aps(ret, rets, '<Default Extension="png"  ContentType="image/png"                                                />');
    aps(ret, rets, '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />');
--    ap(ret, '<Default Extension="xml"  ContentType="application/xml"                                          />');

    if xlsx.content_type_vmlDrawing then
      aps(ret, rets, '<Default Extension="vml"  ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing" />'); -- Needed for Drawings (of which check boxes are one)
    end if;

    aps(ret, rets, '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />');

    for s in 1 ..xlsx.sheets.count loop
      aps(ret, rets, '<Override PartName="/xl/worksheets/sheet' || s || '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />');
    end loop;

--    ap(ret, '<Override PartName="/xl/theme/theme1.xml"      ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />');
    aps(ret, rets, '<Override PartName="/xl/styles.xml"            ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"        />');
    aps(ret, rets, '<Override PartName="/xl/sharedStrings.xml"     ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />');
    aps(ret, rets, '<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"                     />');
    aps(ret, rets, '<Override PartName="/xl/calcChain.xml"         ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"     />');
    aps(ret, rets, '<Override PartName="/docProps/core.xml"        ContentType="application/vnd.openxmlformats-package.core-properties+xml"                    />');
    aps(ret, rets, '<Override PartName="/docProps/app.xml"         ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"         />');

    aps(ret, rets, '</Types>');

    ap(ret,rets);

    return ret;

  end content_types; -- }}}

  function xl_calcChain(xlsx in out book_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000) := null;
  begin

    aps(ret, rets, '<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');

    for e in 1 .. xlsx.calc_chain_elems.count loop

        aps(ret, rets, '<c');
        add_attrs(ret, rets, 'r', xlsx.calc_chain_elems(e).cell_reference);
        add_attrs(ret, rets, 'i', xlsx.calc_chain_elems(e).sheet         );
        aps(ret, rets, '/>');

    end loop;

    aps(ret, rets, '</calcChain>');
    ap(ret,rets);

    return ret;

  end xl_calcChain; -- }}}

  function vml_drawing(v vml_drawing_r) return blob is -- {{{
    ret blob := start_xml_blob;
    rets varchar2(4000) := null;

  begin

    aps(ret, rets, '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">');

    for cb in 1 .. v.checkboxes.count loop

      aps(ret, rets, '<v:shape
         type="#_x0000_t201" 
         filled="f"
         fillcolor="window [65]"
         stroked="f" 
         strokecolor="windowText [64]"
         o:insetmode="auto"
       >');

      aps(ret, rets, '<v:textbox style="mso-direction-alt:auto" o:singleclick="f">
      <div style="text-align:left">
        <font face="Tahoma" size="160" color="auto">' || v.checkboxes(cb).text || '</font>
      </div>
    </v:textbox>');

    aps(ret, rets, '<x:ClientData ObjectType="Checkbox">
      <x:SizeWithCells />
      <x:Anchor>' ||
        (v.checkboxes(cb).col_left -1) || ',' || -- Left column
        '0,'                                  || -- Left offset
        (v.checkboxes(cb).row_top  -1) || ',' || -- Top row
        '0,'                                  || -- Top offset
        (v.checkboxes(cb).col_left -1) || ',' || -- Right column
        '103,'                                || -- Right Offset
        (v.checkboxes(cb).row_top  -1) || ',' || -- Bottom row
        '17                                   || -- Bottom offset
      </x:Anchor> 
      <x:AutoFill>False</x:AutoFill>
      <x:AutoLine>False</x:AutoLine>
      <x:TextVAlign>Center</x:TextVAlign>
      <x:Checked>' || case when v.checkboxes(cb).checked then 1 else 0 end || '</x:Checked>
      <x:NoThreeD />
    </x:ClientData>');


      aps(ret, rets, '</v:shape>');

    end loop;

    aps(ret, rets, '</xml>');
    ap(ret,rets);

    return ret;

  end vml_drawing; -- }}}

  -- }}}

  function create_xlsx(xlsx in out book_r) return blob is -- {{{

    xlsx_b   blob;

    xb_xl_worksheets_rels_sheet   blob;

    procedure add_blob_to_zip(zip      in out blob, -- {{{
                              filename        varchar2,
                              b               blob) is

      b_ blob;
    begin

      b_ := b;

      zipper.addFile(zip, filename, b_);

      dbms_lob.freeTemporary(b_);


    end add_blob_to_zip; -- }}}

  begin

    dbms_lob.createTemporary(xlsx_b, true);

    for s in 1 .. xlsx.sheets.count loop
       xb_xl_worksheets_rels_sheet  := xl_worksheets_rels_sheet  (xlsx, s);
       if xb_xl_worksheets_rels_sheet is not null then
          zipper.addFile(xlsx_b, 'xl/worksheets/_rels/sheet' || s || '.xml.rels', xb_xl_worksheets_rels_sheet);
       end if;
    end loop;

    for s in 1 .. xlsx.sheets.count loop
        zipper.addFile(xlsx_b, 'xl/worksheets/sheet' || s || '.xml', xl_worksheets_sheet       (xlsx, s));

        if xlsx.sheets(s).vml_drawings is not null then

        -- Assumption ASSMPT_01: at most ONE vml drawing per sheet.

           add_blob_to_zip(xlsx_b, 'xl/drawings/vmlDrawing' || s || '.vml', vml_drawing(xlsx.sheets(s).vml_drawings(1)));

        end if;

    end loop;

    add_blob_to_zip(xlsx_b, '_rels/.rels'                        , rels_rels                 (xlsx));
    add_blob_to_zip(xlsx_b, 'docProps/app.xml'                   , docProps_app              ()    );
    add_blob_to_zip(xlsx_b, 'docProps/core.xml'                  , docProps_core             (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/_rels/workbook.xml.rels'         , xl_rels_workbook          (xlsx));

    if xlsx.calc_chain_elems.count > 0 then
       add_blob_to_zip(xlsx_b, 'xl/calcChain.xml'                , xl_calcChain              (xlsx));
    end if;

    add_blob_to_zip(xlsx_b, 'xl/drawings/_rels/drawing1.xml.rels', xl_drawings_rels_drawing1 (xlsx));

    for d in 1 .. xlsx.drawings.count loop
      add_blob_to_zip(xlsx_b, 'xl/drawings/drawing' || d || '.xml',  utl_raw.cast_to_raw(xlsx.drawings(d).raw_));
    end loop;

    for m in 1 .. xlsx.medias.count loop
      add_blob_to_zip(xlsx_b, 'xl/media/' || xlsx.medias(m).name_, xlsx.medias(m).b);
    end loop;

    add_blob_to_zip(xlsx_b, 'xl/sharedStrings.xml'               , xl_sharedStrings          (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/styles.xml'                      , xl_styles                 (xlsx));
    add_blob_to_zip(xlsx_b, 'xl/workbook.xml'                    , xl_workbook               (xlsx));
    add_blob_to_zip(xlsx_b, '[Content_Types].xml'                , content_types             (xlsx));

    zipper.finish(xlsx_b);

    return xlsx_b;

--exception when others then
--   raise_application_error(-20800, 'xlsx_writer.create_xlsx, step: ' || step || ', ' || sqlerrm);
  end create_xlsx; -- }}}

  function sql_to_xlsx(sql_stmt varchar2) return blob is -- {{{
 -- {{{
    workbook xlsx_writer.book_r;
    sheet        integer;
    cs_date      integer;

    xlsx         blob;

    cursor_             integer;
    res_                integer;
    column_count        integer;
    column_value        varchar2(4000);
    table_desc_         dbms_sql.desc_tab;
    type column_t       is record(name varchar2(30), datatype char(1) /* N, D, C */, max_characters number);
    type columns_t      is table of column_t;
    column_  column_t;
    columns_  columns_t := columns_t();
 -- }}}

    procedure column_names_and_types is -- {{{
    begin
  
        for c in 1 .. column_count loop -- {{{
  
            column_.name         :=  table_desc_(c).col_name;
  
            column_.datatype     :=  case table_desc_(c).col_type 
                                     when dbms_sql.number_type   then 'N'
                                     when dbms_sql.date_type     then 'D'
                                     when dbms_sql.varchar2_type then 'C'
                                     when dbms_sql.char_type     then 'C'
                                     else '??' -- does not fit into char(1), abort!
                                     end;
  
            columns_.extend;
            columns_(c) := column_;
  
        end loop; -- }}}
  
    end column_names_and_types; -- }}}
    procedure header is -- {{{
    begin

      for c in 1 .. column_count loop -- {{{
          add_cell(workbook, sheet, 1, c, text => columns_(c).name);

          if columns_(c).datatype = 'D' then
             col_width(workbook, sheet, c, 17);
          end if;

      end loop; -- }}}

    end header; -- }}}
    procedure result_set is -- {{{

      cur_row integer := 1;
    begin

        loop -- {{{

            exit when dbms_sql.fetch_rows(cursor_) = 0;

            cur_row := 1 + cur_row;

            for c in 1 .. column_count loop

                dbms_sql.column_value(cursor_, c, column_value);

                if     columns_(c).datatype = 'N' then
                       add_cell(workbook, sheet, cur_row, c, value_ => column_value);

                elsif  columns_(c).datatype = 'D' then
                       add_cell(workbook, sheet, cur_row, c, date_  => column_value, style_id => cs_date);

                elsif  columns_(c).datatype = 'C' then
                       add_cell(workbook, sheet, cur_row, c, text   => column_value);

                       columns_(c).max_characters := greatest(nvl(columns_(c).max_characters, 0), nvl(length(column_value), 0));
                end if;

            end loop;


        end loop; -- }}}

    end result_set; -- }}}

  begin

    workbook := xlsx_writer.start_book;
    sheet    := xlsx_writer.add_sheet  (workbook, 'Result set');

    cs_date := xlsx_writer.add_cell_style(workbook, num_fmt_id => "m/d/yy h:mm");

    cursor_  := dbms_sql.open_cursor;
    dbms_sql.parse(cursor_, sql_stmt, dbms_sql.native);
    dbms_sql.describe_columns(/*in*/ cursor_, /*out*/ column_count, /*out*/ table_desc_);
    for c in 1 .. column_count loop -- {
        dbms_sql.define_column(cursor_, c, column_value, 4000);
    end loop; -- }
    res_ := dbms_sql.execute(cursor_);
    column_names_and_types;
    header;
    result_set;
    for c in 1 .. column_count loop -- {

        if columns_(c).datatype = 'C' then

           if columns_(c).max_characters > 13 then
               col_width(workbook, sheet, c, least(columns_(c).max_characters, 50) * 0.95);
           end if;

        end if;
    end loop;

    xlsx     := xlsx_writer.create_xlsx(workbook);

    return xlsx;

  end sql_to_xlsx; -- }}}

  begin
    dbms_session.set_nls('nls_numeric_characters', '''. ''');

end xlsx_writer; -- }}}
/
show errors