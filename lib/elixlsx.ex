defmodule Workbook do
  defstruct sheets: [], datetime: nil
  @type t :: %Workbook{
      sheets: nonempty_list(Sheet.t),
      datetime: Elixlsx.Util.calendar
  }
end

defmodule StringDB do
  @moduledoc ~S"""
  Strings in XLSX can be stored in a sharedStrings.xml file and be looked up
  by ID. This module handles collection of the data in the preprocessing phase.
  """
  defstruct strings: %{}, element_count: 0

  @type t :: %StringDB {
    strings: %{String.t => non_neg_integer},
    element_count: non_neg_integer
  }

  @spec register_string(StringDB.t, String.t) :: StringDB.t
  def register_string(stringdb, s) do
    case Dict.fetch(stringdb.strings, s) do
      :error -> %StringDB{strings: Dict.put(stringdb.strings, s, stringdb.element_count),
                          element_count: stringdb.element_count + 1}
      {:ok, _} -> stringdb
    end
  end

  def get_id(stringdb, s) do
    case Dict.fetch(stringdb.strings, s) do
      :error ->
        raise %ArgumentError{
          message: "Invalid key provided for StringDB.get_id: " <> inspect(s)}
      {:ok, id} ->
        id
    end
  end

  def sorted_id_string_tuples(stringdb) do
    Enum.map(stringdb.strings, fn ({k, v}) -> {v, k} end) |> Enum.sort
  end
end

defmodule Font do
  defstruct bold: false

  def from_props props do
    %Font{bold: !!props[:bold]}
  end
end


defmodule FontDB do
  defstruct fonts: %{}, element_count: 0

  @type t :: %FontDB {
    fonts: %{Font.t => non_neg_integer},
    element_count: non_neg_integer
  }

  @spec register_font(FontDB.t, Font.t) :: FontDB.t
  def register_font(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error -> %FontDB{fonts: Dict.put(fontdb.fonts, font, fontdb.element_count),
                       element_count: fontdb.element_count + 1}
      {:ok, _} -> fontdb
    end
  end

  def get_id(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for FontDB.get_font: " <> inspect(font)}
      {:ok, id} ->
        id
    end
  end

  def sorted_id_font_tuples(fontdb) do
    Enum.map(fontdb.fonts, fn ({k, v}) -> {v, k} end) |> Enum.sort
  end
end


defmodule CellStyle do
  defstruct font: nil

  @type t :: %CellStyle{
    font: Font.t
  }


  def from_props props do
    font = Font.from_props props
    %CellStyle{font: font}
  end
end


defmodule CellStyleDB do
  defstruct cellstyles: %{}, element_count: 0

  @type t :: %CellStyleDB {
    cellstyles: %{CellStyle.t => non_neg_integer},
    element_count: non_neg_integer
  }


  def register_style(cellstyledb, style) do
    case Dict.fetch(cellstyledb.cellstyles, style) do
      :error ->
        # add +1 here already since "0" refers to the default style
        csdb = update_in cellstyledb.cellstyles,
                  &(Dict.put &1, style, (cellstyledb.element_count + 1))
        update_in csdb.element_count, &(&1 + 1)
      {:ok, _} ->
        cellstyledb
    end
  end

  def get_id(cellstyledb, style) do
    case Dict.fetch(cellstyledb.cellstyles, style) do
      :error ->
        raise %ArgumentError{message: "Could not find key in styledb: " <> inspect(style)}
      {:ok, key} ->
        key
    end
  end

  @doc ~S"""
  Recursively register all fonts, border* and fill* properties (*=TBD)
  in the WorkbookCompInfo structure.
  """
  @spec register_all(WorkbookCompInfo.t) :: WorkbookCompInfo.t
  def register_all(wci) do
    Enum.reduce wci.cellstyledb.cellstyles, wci, fn ({style, _}, wci) ->
      update_in wci.fontdb, &(FontDB.register_font &1, style.font)
      # TODO: update_in wci.borderstyledb ...; wci.fillstyledb...
    end
  end
end


defmodule SheetCompInfo do
  @moduledoc ~S"""
  Compilation info for a sheet, to be filled during the actual
  write process.
  """
  defstruct rId: "", filename: "sheet1.xml", sheetId: 0
  @type t :: %SheetCompInfo{
    rId: String.t,
    filename: String.t,
    sheetId: non_neg_integer
  }

  @spec make(non_neg_integer, non_neg_integer) :: SheetCompInfo.t
  def make sheetidx, rId do
    %SheetCompInfo{rId: "rId" <> to_string(rId),
                   filename: "sheet" <> to_string(sheetidx) <> ".xml",
                   sheetId: sheetidx}
  end
end


defmodule WorkbookCompInfo do
  @moduledoc ~S"""
  This module aggregates information about the metainformation
  required to generate the XML file.

  It is used as the aggregator when folding over the individual
  cells.
  """
  defstruct sheet_info: nil,
            stringdb: %StringDB{},
            fontdb: %FontDB{},
            cellstyledb: %CellStyleDB{},
            next_free_xl_rid: nil
end


defmodule Sheet do
  defstruct name: "", rows: [], sheetCompInfo: nil
  @type t :: %Sheet {
    name: String.t,
    rows: list(list(any())),
  }
end


defmodule Elixlsx do
  alias Elixlsx.Util, as: U
  alias Elixlsx.XMLTemplates

  @doc ~S"""
  Accepts a list of Sheets and the next free relationship ID.
  Returns a tuple containing a list of SheetCompInfo's and the next free
  relationship ID.
  """
  @spec make_sheet_info(nonempty_list(Sheet.t), non_neg_integer) :: {list(SheetCompInfo.t), non_neg_integer}
  def make_sheet_info(sheets, init_rId) do
    # fold helper. aggregator holds {list(SheetCompInfo), sheetidx, rId}.
    add_sheet =
      fn (_, {sci, idx, rId}) ->
        {[SheetCompInfo.make(idx, rId) | sci], idx + 1, rId + 1}
      end

    # TODO probably better to use a zip [1..] |> map instead of fold[l|r]/reverse
    {sheetCompInfos, _, nextrID} = List.foldl(sheets, {[], 1, init_rId}, add_sheet)
    {Enum.reverse(sheetCompInfos), nextrID}
  end

  def compinfo_cell_pass_value wci, value do
    cond do
      is_binary(value) && String.valid?(value)
        -> update_in wci.stringdb, &StringDB.register_string(&1, value)
      true -> wci
    end
  end


  def compinfo_cell_pass_style wci, props do
    update_in wci.cellstyledb, &CellStyleDB.register_style(&1, CellStyle.from_props(props))
  end


  @spec compinfo_cell_pass(WorkbookCompInfo.t, any) :: WorkbookCompInfo.t
  def compinfo_cell_pass wci, cell do
    cond do
      is_list(cell) ->
        wci
        |> compinfo_cell_pass_value(hd cell)
        |> compinfo_cell_pass_style(tl cell)
      true ->
        compinfo_cell_pass_value wci, cell
    end
  end


  @spec compinfo_from_rows(WorkbookCompInfo.t, list(list(any()))) :: WorkbookCompInfo.t
  def compinfo_from_rows wci, rows do
    List.foldl rows, wci, fn (cols, wci) ->
      List.foldl cols, wci, fn (cell, wci) ->
        compinfo_cell_pass wci, cell
      end
    end
  end

  @spec compinfo_from_sheets(WorkbookCompInfo.t, list(Sheet.t)) :: WorkbookCompInfo.t
  def compinfo_from_sheets wci, sheets do
    List.foldl sheets, wci, fn (sheet, wci) ->
      compinfo_from_rows wci, sheet.rows
    end
  end

  @first_free_rid 2
  def make_workbook_comp_info workbook do
    {sci, next_rId} = make_sheet_info(workbook.sheets, @first_free_rid)

    %WorkbookCompInfo{
      sheet_info: sci,
      next_free_xl_rid: next_rId,
    }
    |> compinfo_from_sheets(workbook.sheets)
    |> CellStyleDB.register_all
  end


  @doc ~S"""
  returns a tuple {'docProps/app.xml', "XML Data"}
  """
  def get_docProps_app_xml(data) do
    {'docProps/app.xml', XMLTemplates.docprops_app}
  end


  @spec get_docProps_core_xml(Workbook.t) :: String.t
  def get_docProps_core_xml(workbook) do
    timestamp = U.iso_timestamp(workbook.datetime)
    {'docProps/core.xml', XMLTemplates.docprops_core(timestamp)}
  end


	def get_docProps_dir(data) do
		[get_docProps_app_xml(data), get_docProps_core_xml(data)]
	end


	def get__rels_dotrels(data) do
		{'_rels/.rels',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
		}
	end

	def get__rels_dir(data) do
		[get__rels_dotrels(data)]
	end

	def get_xl_rels_dir(data, sheetCompInfos, next_rId) do
		[{'xl/_rels/workbook.xml.rels',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  """
  <>
  XMLTemplates.make_xl_rel_sheets(sheetCompInfos)
  <>
  """
  <Relationship Id="rId#{next_rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
"""
		}]
	end

	def get_xl_styles_xml(data) do
		{'xl/styles.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="GENERAL"/>
  </numFmts>
  <fonts count="1">
    <font>
      <sz val="10"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle builtinId="0" name="Normal" xfId="0"/>
  </cellStyles>
</styleSheet>
"""
		}
	end

	def get_xl_workbook_xml(data, sheetCompInfos) do
		{'xl/workbook.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="Calc"/>
  <bookViews>
    <workbookView activeTab="0"/>
  </bookViews>
  <sheets>
  """ 
  <> XMLTemplates.make_xl_workbook_xml_sheet_entries(data.sheets, sheetCompInfos)
  <> ~S"""
  </sheets>
  <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
</workbook>
"""
		}
	end

	def get_xl_sharedStrings_xml(data, wci) do
    {'xl/sharedStrings.xml',
			XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples wci.stringdb)
		}
	end


  @spec sheet_full_path(SheetCompInfo.t) :: list(char)
  defp sheet_full_path sci do
    String.to_char_list "xl/worksheets/#{sci.filename}"
  end


	def get_xl_worksheets_dir(data, wci) do
    sheets = data.sheets
    Enum.zip(sheets, wci.sheet_info)
    |> Enum.map fn ({s, sci}) ->
                  {sheet_full_path(sci), XMLTemplates.make_sheet(s, wci)}
                end
	end


	def get_contentTypes_xml(data, wci) do
		{'[Content_Types].xml',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  """ <> XMLTemplates.make_content_types_xml_sheet_entries(wci.sheet_info) <>
  ~S"""
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"""
		}
	end


	def get_xl_dir(data, wci) do
    sheet_comp_infos = wci.sheet_info
    next_free_xl_rid = wci.next_free_xl_rid

		[ get_xl_styles_xml(data),
			get_xl_sharedStrings_xml(data, wci),
			get_xl_workbook_xml(data, sheet_comp_infos)] ++
		get_xl_rels_dir(data, sheet_comp_infos, next_free_xl_rid) ++
		get_xl_worksheets_dir(data, wci)
	end


  def write_to(workbook, filename) do
    wci = make_workbook_comp_info workbook
    IO.inspect wci
    :zip.create(filename,
			get_docProps_dir(workbook) ++
      get__rels_dir(workbook) ++
      get_xl_dir(workbook, wci) ++
      [ get_contentTypes_xml(workbook, wci) ])
  end
end

