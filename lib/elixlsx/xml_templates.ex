defmodule Elixlsx.XMLTemplates do
  alias Elixlsx.Util, as: U
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.FontDB
  alias Elixlsx.Compiler.StringDB

  @doc ~S"""
  There are 5 characters that should be escaped in XML (<,>,",',&), but only
  2 of them *must* be escaped. Saves a couple of CPU cycles, for the environment.

  Example:
    iex> Elixlsx.XMLTemplates.minimal_xml_text_escape "Only '&' and '<' are escaped here, '\"' & '>' & \"'\" are not."
    "Only '&amp;' and '&lt;' are escaped here, '\"' &amp; '>' &amp; \"'\" are not."

  """
  def minimal_xml_text_escape(s) do
    s |> String.replace("&", "&amp;") |> String.replace("<", "&lt;")
  end

  @docprops_app ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Elixlsx</Application>
  <AppVersion>0.0.1</AppVersion>
</Properties>
"""
  def docprops_app, do: @docprops_app


  @docprops_core ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">__TIMESTAMP__</dcterms:created>
  <dc:language>__LANGUAGE__</dc:language>
  <dcterms:modified xsi:type="dcterms:W3CDTF">__TIMESTAMP__</dcterms:modified>
  <cp:revision>__REVISION__</cp:revision>
</cp:coreProperties>
"""

  def docprops_core(timestamp, language \\ "en-US", revision \\ 1) do
    @docprops_core |>
    String.replace("__TIMESTAMP__", timestamp) |>
    String.replace("__LANGUAGE__", language) |>
    String.replace("__REVISION__", to_string(revision))
  end 


  @spec make_xl_rel_sheet(SheetCompInfo.t) :: String.t
  def make_xl_rel_sheet sheet_comp_info do
    # I'd love to use string interpolation here, but unfortunately """< is heredoc notation, so i have to use
    # string concatenation or escape all the quotes. Choosing the first.
    "<Relationship Id=\"#{sheet_comp_info.rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/#{sheet_comp_info.filename}\"/>"
  end


  @spec make_xl_rel_sheets(nonempty_list(SheetCompInfo.t)) :: String.t
  def make_xl_rel_sheets sheet_comp_infos do
    Enum.map_join sheet_comp_infos, &make_xl_rel_sheet/1
  end

  ### xl/workbook.xml
  @spec make_xl_workbook_xml_sheet_entry({Sheet.t, SheetCompInfo.t}) :: String.t
  def make_xl_workbook_xml_sheet_entry {sheet_info, sheet_comp_info} do
    """
<sheet name="#{sheet_info.name}" sheetId="#{sheet_comp_info.sheetId}" state="visible" r:id="#{sheet_comp_info.rId}"/>
    """
  end
  
  @spec make_xl_workbook_xml_sheet_entries(nonempty_list(Sheet.t), nonempty_list(SheetCompInfo.t)) :: String.t
  def make_xl_workbook_xml_sheet_entries sheet_infos, sheet_comp_infos do
    # Yes, the parens here are 100% required, otherwise |> takes precedence,
    # and you'll be spending quite a while debugging why there is no matching
    # function clause for make_xl_workbook_xml_sheet_entry.
    Enum.zip(sheet_infos, sheet_comp_infos)
    |> Enum.map_join &make_xl_workbook_xml_sheet_entry/1
  end

  ### [Content_Types].xml
  def make_content_types_xml_sheet_entry sheet_comp_info do
    """
    <Override PartName="/xl/worksheets/#{sheet_comp_info.filename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    """
  end

  def make_content_types_xml_sheet_entries sheet_comp_infos do
    Enum.map_join sheet_comp_infos, &make_content_types_xml_sheet_entry/1
  end

  defp split_into_content_style(cell, wci) do
    cond do
      is_list(cell) ->
        {
          hd(cell),
          CellStyleDB.get_id(wci.cellstyledb, (CellStyle.from_props (tl cell)))
        }
      true ->
        {
          cell,
          0
        }
    end
  end

  defp get_content_type_value(content, wci) do
    cond do
      is_number content ->
        {"n", to_string(content)}
      String.valid? content ->
        {"s", to_string(StringDB.get_id wci.stringdb, content)}
      true ->
        :error
    end
  end

  ### Worksheets
  # TODO i know now about string interpolation, i should probably clean this up. ;)
  defp xl_sheet_cols(row, rowidx, wci) do
    Enum.zip(row, 1 .. length row) |>
    Enum.map(
      fn {cell, colidx} ->
        {content, styleID} = split_into_content_style(cell, wci)
        cv = get_content_type_value(content, wci)
        {contentType, contentValue} =
          case cv do
            {t, v} -> {t, v}
            :error -> raise %ArgumentError{
                           message: "Invalid column content at " <>
                                    U.to_excel_coords(rowidx, colidx) <> ": "
                                    <> (inspect content)
                                     }
          end

        """
        <c r="#{U.to_excel_coords(rowidx, colidx)}"
           s="#{styleID}" t="#{contentType}">
          <v>#{contentValue}</v>
        </c>
        """
        end) |>
    List.foldr "", &<>/2
  end


  defp xl_sheet_rows(data, wci) do
    Enum.zip(data, 1 .. length data) |>
    Enum.map(fn {row, rowidx} -> 
        List.foldr(["<row r=\"",
                     to_string(rowidx),
                     "\">\n",
                     xl_sheet_cols(row, rowidx, wci),
                     "</row>"], "", &<>/2) end) |>
    List.foldr("", &<>/2)
  end

  @doc ~S"""
  Returns the complete sheet XML data.
  """
  def make_sheet(sheet, wci) do
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr filterMode="false">
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="12.8"/>
  <sheetData>
  """ 
  <>
  xl_sheet_rows(sheet.rows, wci)
  <>
  ~S"""
  </sheetData>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
</worksheet>
			"""
  end

  @spec make_xl_shared_strings(list({non_neg_integer, String.t})) :: String.t
  def make_xl_shared_strings(stringlist) do
    len = length stringlist
  """
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{len}" uniqueCount="#{len}">
  """
  <> Enum.map_join(stringlist, fn ({_, value}) ->
    # the only two characters that *must* be replaced for safe XML encoding are & and <:
    value_ = (value |> String.replace("&", "&amp;") |> String.replace("<", "&lt;"))
    "<si><t>#{minimal_xml_text_escape value_}</t></si>"
  end)
  <> "</sst>"
  end

  defp font_to_xml_entry(font) do
    bold = if font.bold do "<b val=\"1\"/>" else "" end
    italic = if font.italic do "<i val=\"1\"/>" else "" end
    underline = if font.underline do "<u val=\"1\"/>" else "" end
    strike = if font.strike do "<strike val=\"1\"/>" else "" end
    size = if font.size do
      case is_number(font.size) do
        true ->
          "<sz val=\"#{font.size}\"/>"
        false ->
          raise %ArgumentError{message: "Invalid font size: " <> (inspect font.size)}
      end
    else
      ""
    end

    "<font>#{bold}#{italic}#{underline}#{strike}#{size}</font>"
  end

  def make_font_list(ordered_font_list) do
    Enum.map_join(ordered_font_list, "\n", &(font_to_xml_entry &1))
  end

  defp style_to_xml_entry(style, wci) do
    fontid = Elixlsx.Compiler.FontDB.get_id wci.fontdb, style.font
    """
    <xf borderId="0"
           fillId="0"
           fontId=\"#{fontid}\"
           numFmtId="0"
           xfId="0" />
    """
  end

  def make_cellxfs(ordered_style_list, wci) do
    Enum.map_join(ordered_style_list, "\n", &(style_to_xml_entry &1, wci))
  end

  def make_xl_styles(wci) do
    font_list = Elixlsx.Compiler.FontDB.id_sorted_fonts wci.fontdb
    cellXfs = CellStyleDB.id_sorted_styles wci.cellstyledb

		"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="#{1 + length font_list}">
    <font />
    #{make_font_list(font_list)}
  </fonts>
  <fills count="1">
    <fill />
  </fills>
  <borders count="1">
    <border />
  </borders>
  <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0"/>
  </cellStyleXfs>
  <cellXfs count="#{1 + length cellXfs}">
    <xf borderId="0" fillId="0" fontId="0" xfId="0"/>
    #{make_cellxfs cellXfs, wci}
  </cellXfs>
  </styleSheet>
  """

#  <cellStyles count="1">
#    <cellStyle builtinId="0" name="Normal" xfId="0"/>
#  </cellStyles>
	end
end
