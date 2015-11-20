defmodule Elixlsx.XMLTemplates do
  alias Elixlsx.Util, as: U

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


  ### Worksheets

  defp xl_sheet_cols(row, rowidx) do
    Enum.zip(row, 1 .. length row) |>
    Enum.map(
      fn {col, colidx} ->
        {type, value} = cond do
          is_number(col) -> {"n", to_string(col)}
          true -> {"s", to_string(col)} # TODO this may throw.
        end
        List.foldr ["<c r=\"",
                     U.to_excel_coords(rowidx, colidx),
                     "\" s=\"0\" t=\"",
                     type,
                     "\">",
                     "<v>",
                     value,
                     "</v></c>"], "", &<>/2
        end) |>
    List.foldr "", &<>/2
  end


  defp xl_sheet_rows(data) do
    Enum.zip(data, 1 .. length data) |>
    Enum.map(fn {row, rowidx} -> 
        List.foldr(["<row r=\"",
                     to_string(rowidx),
                     "\">\n",
                     xl_sheet_cols(row, rowidx),
                     "</row>"], "", &<>/2) end) |>
    List.foldr("", &<>/2)
  end

  @doc ~S"""
  Returns the complete sheet XML data.
  """
  def make_sheet sheet do
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
  xl_sheet_rows(sheet.rows)
  <>
  ~S"""
  </sheetData>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
</worksheet>
			"""
  end
end
