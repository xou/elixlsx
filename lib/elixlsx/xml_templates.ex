defmodule Elixlsx.XMLTemplates do
  alias Elixlsx.Util, as: U
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.Compiler.FontDB
  alias Elixlsx.Compiler.FillDB
  alias Elixlsx.Compiler.SheetCompInfo
  alias Elixlsx.Compiler.NumFmtDB
  alias Elixlsx.Compiler.BorderStyleDB
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Style.CellStyle
  alias Elixlsx.Style.Font
  alias Elixlsx.Style.Fill
  alias Elixlsx.Style.BorderStyle
  alias Elixlsx.Sheet

  # TODO: the xml_text_exape functions belong into Elixlsx.Util,
  # as they are/will be used by functions in Elixlsx.Style.*
  @doc ~S"""
  There are 5 characters that should be escaped in XML (<,>,",',&), but only
  2 of them *must* be escaped. Saves a couple of CPU cycles, for the environment.

  ## Examples

      iex> Elixlsx.XMLTemplates.minimal_xml_text_escape "Only '&' and '<' are escaped here, '\"' & '>' & \"'\" are not."
      "Only '&amp;' and '&lt;' are escaped here, '\"' &amp; '>' &amp; \"'\" are not."

  """
  def minimal_xml_text_escape(s) do
    U.replace_all(s, [{"&", "&amp;"}, {"<", "&lt;"}])
  end

  @doc ~S"""
  Escape characters for embedding in XML
  documents.

  ## Examples

      iex> Elixlsx.XMLTemplates.xml_escape "&\"'<>'"
      "&amp;&quot;&apos;&lt;&gt;&apos;"

  """
  def xml_escape(s) do
    U.replace_all(s, [
      {"&", "&amp;"},
      {"'", "&apos;"},
      {"\"", "&quot;"},
      {"<", "&lt;"},
      {">", "&gt;"}
    ])
  end

  @docprops_app ~S"""
  <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <TotalTime>0</TotalTime>
    <Application>Elixlsx</Application>
    <AppVersion>__APPVERSION__</AppVersion>
  </Properties>
  """
  def docprops_app do
    U.replace_all(
      @docprops_app,
      [{"__APPVERSION__", U.app_version_string()}]
    )
  end

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
    U.replace_all(
      @docprops_core,
      [
        {"__TIMESTAMP__", xml_escape(timestamp)},
        {"__LANGUAGE__", language},
        {"__REVISION__", to_string(revision)}
      ]
    )
  end

  @spec make_xl_rel_sheet(SheetCompInfo.t()) :: String.t()
  def make_xl_rel_sheet(sheet_comp_info) do
    # I'd love to use string interpolation here, but unfortunately """< is heredoc notation, so i have to use
    # string concatenation or escape all the quotes. Choosing the first.
    "<Relationship Id=\"#{sheet_comp_info.rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/#{sheet_comp_info.filename}\"/>"
  end

  @spec make_xl_rel_sheets(nonempty_list(SheetCompInfo.t())) :: String.t()
  def make_xl_rel_sheets(sheet_comp_infos) do
    Enum.map_join(sheet_comp_infos, &make_xl_rel_sheet/1)
  end

  ### xl/workbook.xml
  @spec make_xl_workbook_xml_sheet_entry({Sheet.t(), SheetCompInfo.t()}) :: String.t()
  def make_xl_workbook_xml_sheet_entry({sheet_info, sheet_comp_info}) do
    if sheet_info.name == "" do
      raise %ArgumentError{
        message: "The sheet name cannot be blank."
      }
    end

    if String.length(sheet_info.name) > 31 do
      raise %ArgumentError{
        message:
          "The sheet name '#{sheet_info.name}' is too long. Maximum 31 chars allowed for name."
      }
    end

    if String.contains?(sheet_info.name, ~W(: \ / ? * [ ])) do
      raise %ArgumentError{
        message:
          "The sheet name '#{sheet_info.name}' contains following invalid characters: : \ / ? * [ ])"
      }
    end

    """
    <sheet name="#{xml_escape(sheet_info.name)}" sheetId="#{sheet_comp_info.sheetId}" state="visible" r:id="#{sheet_comp_info.rId}"/>
    """
  end

  ### [Content_Types].xml
  defp contenttypes_sheet_entry(sheet_comp_info) do
    """
    <Override PartName="/xl/worksheets/#{sheet_comp_info.filename}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
    """
  end

  defp contenttypes_sheet_entries(sheet_comp_infos) do
    Enum.map_join(sheet_comp_infos, &contenttypes_sheet_entry/1)
  end

  def make_contenttypes_xml(wci) do
    ~S"""
    <?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    """ <>
      contenttypes_sheet_entries(wci.sheet_info) <>
      ~S"""
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
      </Types>
      """
  end

  ###
  ### xl/worksheet/sheet*.xml
  ###

  defp split_into_content_style([h | t], wci) do
    cellstyle = CellStyle.from_props(t)

    {
      h,
      CellStyleDB.get_id(wci.cellstyledb, cellstyle),
      cellstyle
    }
  end

  defp split_into_content_style(cell, _wci), do: {cell, 0, nil}

  defp get_content_type_value(content, wci) do
    case content do
      {:excelts, num} ->
        {"n", to_string(num)}

      {:formula, x} ->
        {:formula, x}

      {:formula, x, opts} when is_list(opts) ->
        {:formula, x, opts}

      x when is_number(x) ->
        {"n", to_string(x)}

      x when is_binary(x) ->
        id = StringDB.get_id(wci.stringdb, x)

        if id == -1 do
          {:empty, :empty}
        else
          {"s", to_string(id)}
        end

      x when is_boolean(x) ->
        {"b",
         if x do
           "1"
         else
           "0"
         end}

      :empty ->
        {:empty, :empty}

      _ ->
        :error
    end
  end

  # TODO i know now about string interpolation, i should probably clean this up. ;)
  defp xl_sheet_cols(row, rowidx, wci) do
    {updated_row, _id} =
      row
      |> List.foldl({"", 1}, fn cell, {acc, colidx} ->
        {content, styleID, cellstyle} = split_into_content_style(cell, wci)

        if is_nil(content) do
          {acc, colidx + 1}
        else
          content =
            if CellStyle.is_date?(cellstyle) do
              U.to_excel_datetime(content)
            else
              content
            end

          cv = get_content_type_value(content, wci)

          {content_type, content_value, content_opts} =
            case cv do
              {t, v} ->
                {t, v, []}

              {t, v, opts} ->
                {t, v, opts}

              :error ->
                raise %ArgumentError{
                  message:
                    "Invalid column content at " <>
                      U.to_excel_coords(rowidx, colidx) <> ": " <> inspect(content)
                }
            end

          cell_xml =
            case content_type do
              :formula ->
                value =
                  if not is_nil(content_opts[:value]),
                    do: "<v>#{content_opts[:value]}</v>",
                    else: ""

                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}"
                s="#{styleID}">
                <f>#{content_value}</f>
                #{value}
                </c>
                """

              :empty ->
                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}"
                s="#{styleID}">
                </c>
                """

              type ->
                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}"
                s="#{styleID}" t="#{type}">
                <v>#{content_value}</v>
                </c>
                """
            end

          {acc <> cell_xml, colidx + 1}
        end
      end)

    updated_row
  end

  defp make_data_validations([]) do
    ""
  end

  defp make_data_validations(data_validations) do
    """
    <dataValidations count="#{Enum.count(data_validations)}">
      #{Enum.map(data_validations, &make_data_validation/1)}
    </dataValidations>
    """
  end

  defp make_data_validation({start_cell, end_cell, values}) when is_bitstring(values) do
    """
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="#{start_cell}:#{end_cell}">
      <formula1>#{values}</formula1>
    </dataValidation>
    """
  end

  defp make_data_validation({start_cell, end_cell, values}) do
    joined_values =
      values
      |> Enum.join(",")
      |> String.codepoints()
      |> Enum.chunk_every(255)
      |> Enum.join("&quot;&amp;&quot;")

    """
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="#{start_cell}:#{end_cell}">
      <formula1>&quot;#{joined_values}&quot;</formula1>
    </dataValidation>
    """
  end

  defp xl_merge_cells([]) do
    ""
  end

  defp xl_merge_cells(merge_cells) do
    """
    <mergeCells count="#{Enum.count(merge_cells)}">
      #{Enum.map(merge_cells, fn {fromCell, toCell} -> "<mergeCell ref=\"#{fromCell}:#{toCell}\"/>" end)}
    </mergeCells>
    """
  end

  defp xl_sheet_rows(data, row_heights, grouping_info, wci) do
    rows =
      Enum.zip(data, 1..length(data))
      |> Enum.map_join(fn {row, rowidx} ->
        """
        <row r="#{rowidx}" #{get_row_height_attr(row_heights, rowidx)}#{get_row_grouping_attr(grouping_info, rowidx)}>
          #{xl_sheet_cols(row, rowidx, wci)}
        </row>
        """
      end)

    if (length(data) + 1) in grouping_info.collapsed_idxs do
      rows <>
        """
        <row r="#{length(data) + 1}" collapsed="1"></row>
        """
    else
      rows
    end
  end

  defp get_row_height_attr(row_heights, rowidx) do
    row_height = Map.get(row_heights, rowidx)

    if row_height do
      "customHeight=\"1\" ht=\"#{row_height}\""
    else
      ""
    end
  end

  defp get_row_grouping_attr(gr_info, rowidx) do
    outline_level = Map.get(gr_info.outline_lvs, rowidx)

    if(outline_level, do: " outlineLevel=\"#{outline_level}\"", else: "") <>
      if(rowidx in gr_info.hidden_idxs, do: " hidden=\"1\"", else: "") <>
      if rowidx in gr_info.collapsed_idxs, do: " collapsed=\"1\"", else: ""
  end

  @typep grouping_info :: %{
           outline_lvs: %{optional(idx :: pos_integer) => lv :: pos_integer},
           hidden_idxs: MapSet.t(pos_integer),
           collapsed_idxs: MapSet.t(pos_integer)
         }
  @spec get_grouping_info([Sheet.rowcol_group()]) :: grouping_info
  defp get_grouping_info(groups) do
    ranges =
      Enum.map(groups, fn
        {%Range{} = range, _opts} -> range
        %Range{} = range -> range
      end)

    collapsed_ranges =
      groups
      |> Enum.filter(fn
        {%Range{} = _range, opts} -> opts[:collapsed]
        %Range{} = _range -> false
      end)
      |> Enum.map(fn {range, _opts} -> range end)

    # see ECMA Office Open XML Part1, 18.3.1.73 Row -> attributes -> collapsed for examples
    %{
      outline_lvs:
        ranges
        |> Stream.concat()
        |> Enum.group_by(& &1)
        |> Map.new(fn {k, v} -> {k, length(v)} end),
      hidden_idxs: collapsed_ranges |> Stream.concat() |> MapSet.new(),
      collapsed_idxs: collapsed_ranges |> Enum.map(&(&1.last + 1)) |> MapSet.new()
    }
  end

  defp make_col({k, width, outline_level, hidden, collapsed}) do
    width_attr = if width, do: " width=\"#{width}\" customWidth=\"1\"", else: ""
    hidden_attr = if hidden, do: " hidden=\"1\"", else: ""
    outline_level_attr = if outline_level, do: " outlineLevel=\"#{outline_level}\"", else: ""
    collapsed_attr = if collapsed, do: " collapsed=\"1\"", else: ""

    ~c'<col min="#{k}" max="#{k}"#{width_attr}#{hidden_attr}#{outline_level_attr}#{collapsed_attr} />'
  end

  defp make_cols(sheet) do
    grouping_info = get_grouping_info(sheet.group_cols)

    col_indices =
      Stream.concat([
        Map.keys(sheet.col_widths),
        Map.keys(grouping_info.outline_lvs),
        grouping_info.hidden_idxs,
        grouping_info.collapsed_idxs
      ])
      |> Enum.sort()
      |> Enum.dedup()

    unless Enum.empty?(col_indices) do
      cols =
        col_indices
        |> Stream.map(
          &{
            &1,
            Map.get(sheet.col_widths, &1),
            Map.get(grouping_info.outline_lvs, &1),
            &1 in grouping_info.hidden_idxs,
            &1 in grouping_info.collapsed_idxs
          }
        )
        |> Enum.map_join(&make_col/1)

      "<cols>#{cols}</cols>"
    else
      ""
    end
  end

  defp make_max_outline_level_row(row_outline_levels) do
    unless row_outline_levels === %{} do
      max_outline_level_row =
        Map.values(row_outline_levels)
        |> Enum.max()

      " outlineLevelRow=\"#{max_outline_level_row}\""
    else
      ""
    end
  end

  @spec make_sheet(Sheet.t(), WorkbookCompInfo.t()) :: String.t()
  @doc ~S"""
  Returns the XML content for single sheet.
  """
  def make_sheet(sheet, wci) do
    grouping_info = get_grouping_info(sheet.group_rows)

    ~S"""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetPr filterMode="false">
      <pageSetUpPr fitToPage="false"/>
    </sheetPr>
    <dimension ref="A1"/>
    <sheetViews>
    <sheetView workbookViewId="0"
    """ <>
      make_sheet_show_grid(sheet) <>
      """
      >
      """ <>
      make_sheetview(sheet) <>
      """
      </sheetView>
      </sheetViews>
      <sheetFormatPr defaultRowHeight="12.8"
      """ <>
      make_max_outline_level_row(grouping_info.outline_lvs) <>
      """
      />
      """ <>
      make_cols(sheet) <>
      """
      <sheetData>
      """ <>
      xl_sheet_rows(sheet.rows, sheet.row_heights, grouping_info, wci) <>
      ~S"""
      </sheetData>
      """ <>
      xl_merge_cells(sheet.merge_cells) <>
      make_data_validations(sheet.data_validations) <>
      """
      <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
      </worksheet>
      """
  end

  defp make_sheet_show_grid(sheet) do
    show_grid_lines_xml =
      case sheet.show_grid_lines do
        false -> "showGridLines=\"0\" "
        _ -> ""
      end

    show_grid_lines_xml
  end

  defp make_sheetview(sheet) do
    # according to spec:
    # * when only horizontal split is applied we need to use bottomLeft
    # * when only vertical split is applied we need to use topRight
    # * and when both splits is applied, we can use bottomRight
    pane =
      case sheet.pane_freeze do
        {_row_idx, 0} ->
          "bottomLeft"

        {0, _col_idx} ->
          "topRight"

        {col_idx, row_idx} when col_idx > 0 and row_idx > 0 ->
          "bottomRight"

        _any ->
          nil
      end

    {selection_pane_attr, panel_xml} =
      case sheet.pane_freeze do
        {row_idx, col_idx} when col_idx > 0 or row_idx > 0 ->
          top_left_cell = U.to_excel_coords(row_idx + 1, col_idx + 1)

          {"pane=\"#{pane}\"",
           "<pane xSplit=\"#{col_idx}\" ySplit=\"#{row_idx}\" topLeftCell=\"#{top_left_cell}\" activePane=\"#{pane}\" state=\"frozen\" />"}

        _any ->
          {"", ""}
      end

    panel_xml <> "<selection " <> selection_pane_attr <> " activeCell=\"A1\" sqref=\"A1\" />"
  end

  ###
  ### xl/sharedStrings.xml
  ###

  @spec make_xl_shared_strings(list({non_neg_integer, String.t()})) :: String.t()
  def make_xl_shared_strings(stringlist) do
    len = length(stringlist)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{len}" uniqueCount="#{len}">
    """ <>
      Enum.map_join(stringlist, fn {_, value} ->
        # the only two characters that *must* be replaced for safe XML encoding are & and <:
        "<si><t xml:space=\"preserve\">#{minimal_xml_text_escape(value)}</t></si>"
      end) <> "</sst>"
  end

  ###
  ### xl/styles.xml
  ###

  @spec make_font_list(list(Font.t())) :: String.t()
  defp make_font_list(ordered_font_list) do
    Enum.map_join(ordered_font_list, "\n", &Font.get_stylexml_entry(&1))
  end

  @spec make_fill_list(list(Fill.t())) :: String.t()
  defp make_fill_list(ordered_fill_list) do
    Enum.map_join(ordered_fill_list, "\n", &Fill.get_stylexml_entry(&1))
  end

  # Turns a CellStyle struct into the styles.xml <xf /> representation.
  # TODO: This could be moved into the CellStyle struct.
  @spec style_to_xml_entry(CellStyle.t(), WorkbookCompInfo.t()) :: String.t()
  defp style_to_xml_entry(style, wci) do
    fontid =
      if is_nil(style.font),
        do: 0,
        else: FontDB.get_id(wci.fontdb, style.font)

    fillid =
      if is_nil(style.fill),
        do: 0,
        else: FillDB.get_id(wci.filldb, style.fill)

    numfmtid =
      if is_nil(style.numfmt),
        do: 0,
        else: NumFmtDB.get_id(wci.numfmtdb, style.numfmt)

    borderid =
      if is_nil(style.border),
        do: 0,
        else: BorderStyleDB.get_id(wci.borderstyledb, style.border)

    {apply_alignment, wrap_text_tag} =
      case style.font do
        nil ->
          {"", ""}

        font ->
          case make_style_alignment(font) do
            "" ->
              {"", ""}

            alignment ->
              {"applyAlignment=\"1\"", alignment}
          end
      end

    """
    <xf borderId="#{borderid}"
           fillId="#{fillid}"
           fontId="#{fontid}"
           numFmtId="#{numfmtid}"
           xfId="0" #{apply_alignment}>
      #{wrap_text_tag}
    </xf>
    """
  end

  @spec wrap_text(String.t(), Font.t()) :: String.t()
  defp wrap_text(attrs, %Font{wrap_text: true}), do: attrs <> "wrapText=\"1\" "
  defp wrap_text(attrs, _), do: attrs

  @spec horizontal_alignment(String.t(), Font.t()) :: String.t()
  defp horizontal_alignment(attrs, %Font{align_horizontal: nil}), do: attrs

  defp horizontal_alignment(attrs, %Font{align_horizontal: alignment}) do
    if alignment in [:center, :fill, :general, :justify, :left, :right] do
      attrs <> "horizontal=\"#{Atom.to_string(alignment)}\" "
    else
      raise %ArgumentError{
        message:
          "Given horizontal alignment not supported. Only :center, :fill, :general, :justify, :left, :right are available."
      }
    end
  end

  @spec vertical_alignment(String.t(), Font.t()) :: String.t()
  defp vertical_alignment(attrs, %Font{align_vertical: nil}), do: attrs

  defp vertical_alignment(attrs, %Font{align_vertical: alignment}) do
    if alignment in [:center, :top, :bottom] do
      attrs <> "vertical=\"#{Atom.to_string(alignment)}\" "
    else
      raise %ArgumentError{
        message:
          "Given vertical alignment not supported. Only :center, :top, :bottom are available."
      }
    end
  end

  # Creates an alignment xml tag from font style.
  @spec make_style_alignment(Font.t()) :: String.t()
  defp make_style_alignment(font) do
    attrs =
      ""
      |> wrap_text(font)
      |> horizontal_alignment(font)
      |> vertical_alignment(font)

    case attrs do
      "" ->
        nil

      ^attrs ->
        "<alignment #{attrs}/>"
    end
  end

  # Returns the inner content of the <CellXfs> block.
  @spec make_cellxfs(list(CellStyle.t()), WorkbookCompInfo.t()) :: String.t()
  defp make_cellxfs(ordered_style_list, wci) do
    Enum.map_join(ordered_style_list, "\n", &style_to_xml_entry(&1, wci))
  end

  alias Elixlsx.Style.NumFmt

  defp make_numfmts_inner(id_numfmt_tuples) do
    Enum.map_join(id_numfmt_tuples, "\n", fn {id, numfmt} ->
      NumFmt.get_stylexml_entry(numfmt, id)
    end)
  end

  defp make_numfmts(id_numfmt_tuples) do
    case length(id_numfmt_tuples) do
      0 -> ""
      n -> "<numFmts count=\"#{n}\">#{make_numfmts_inner(id_numfmt_tuples)}</numFmts>"
    end
  end

  defp make_borders(borders_list) do
    Enum.map_join(borders_list, "\n", &BorderStyle.get_border_style_entry(&1))
  end

  @spec make_xl_styles(WorkbookCompInfo.t()) :: String.t()
  @doc ~S"""
  Get the content of the `styles.xml` file.

  The WorkbookCompInfo struct must be computed before calling this,
  (especially CellStyleDB.register_all)
  """
  def make_xl_styles(wci) do
    font_list = FontDB.id_sorted_fonts(wci.fontdb)
    fill_list = FillDB.id_sorted_fills(wci.filldb)
    cell_xfs = CellStyleDB.id_sorted_styles(wci.cellstyledb)
    numfmts_list = NumFmtDB.custom_numfmt_id_tuples(wci.numfmtdb)
    borders_list = BorderStyleDB.id_sorted_borders(wci.borderstyledb)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      #{make_numfmts(numfmts_list)}
      <fonts count="#{1 + length(font_list)}">
        <font />
        #{make_font_list(font_list)}
      </fonts>
      <fills count="#{2 + length(fill_list)}">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        #{make_fill_list(fill_list)}
      </fills>
      <borders count="#{1 + length(borders_list)}">
        <border />
        #{make_borders(borders_list)}
      </borders>
      <cellStyleXfs count="1">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" applyAlignment="1">
          <alignment wrapText="1"/>
        </xf>
      </cellStyleXfs>
      <cellXfs count="#{1 + length(cell_xfs)}">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" xfId="0"/>
        #{make_cellxfs(cell_xfs, wci)}
      </cellXfs>
    </styleSheet>
    """
  end

  ###
  ### _rels/.rels
  ###

  @rels_dotrels ~S"""
  <?xml version="1.0" encoding="UTF-8"?>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
  </Relationships>
  """
  def rels_dotrels, do: @rels_dotrels

  ####
  #### xl/workbook.xml
  ####

  @spec workbook_sheet_entries(nonempty_list(Sheet.t()), nonempty_list(SheetCompInfo.t())) ::
          String.t()
  defp workbook_sheet_entries(sheet_infos, sheet_comp_infos) do
    Enum.zip(sheet_infos, sheet_comp_infos)
    |> Enum.map_join(&make_xl_workbook_xml_sheet_entry/1)
  end

  defp workbook_defined_names(%{}) do
    ""
  end

  defp workbook_defined_names(defined_names) do
    """
    <definedNames>
    #{Enum.map_join(defined_names, &make_defined_name/1)}
    </definedNames>
    """
  end

  defp make_defined_name({name, definition}) do
    """
    <definedName name="#{xml_escape(name)}">#{xml_escape(definition)}</definedName>
    """
  end

  @doc ~S"""
  Return the data for /xl/workbook.xml
  """
  def make_workbook_xml(data, sci) do
    ~S"""
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="Calc"/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
    """ <>
      workbook_sheet_entries(data.sheets, sci) <>
      ~S"""
      </sheets>
      """ <>
      workbook_defined_names(data.defined_names) <>
      """
      <calcPr fullCalcOnLoad="1" iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
      </workbook>
      """
  end
end
