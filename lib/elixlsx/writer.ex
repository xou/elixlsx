defmodule Elixlsx.Writer do
  alias Elixlsx.Util, as: U
  alias Elixlsx.XMLTemplates
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Compiler.SheetCompInfo
  alias Elixlsx.Workbook

  @type zip_tuple :: {charlist, String.t()}

  @moduledoc ~S"""
  Contains functions to generate the individual files
  in the XLSX zip package.
  """

  @spec create_files(Workbook.t(), WorkbookCompInfo.t()) :: list(zip_tuple)
  @doc ~S"""
  Returns a list of tuples {filename, filecontent}. Both
  filename and filecontent are represented as charlists
  (so that they can be used with the OTP :zip module.)
  """
  def create_files(workbook, wci) do
    get_docProps_dir(workbook) ++
      get__rels_dir(workbook) ++
      get_xl_dir(workbook, wci) ++ [get_contentTypes_xml(workbook, wci)]
  end

  @spec get_docProps_app_xml(Workbook.t()) :: zip_tuple
  @doc ~S"""
  returns a tuple {'docProps/app.xml', "XML Data"}
  """
  def get_docProps_app_xml(_) do
    {'docProps/app.xml', XMLTemplates.docprops_app()}
  end

  @spec get_docProps_core_xml(Workbook.t()) :: zip_tuple
  def get_docProps_core_xml(workbook) do
    timestamp = U.iso_timestamp(workbook.datetime)
    {'docProps/core.xml', XMLTemplates.docprops_core(timestamp)}
  end

  @spec get_docProps_dir(Workbook.t()) :: list(zip_tuple)
  @doc ~S"""
  Returns files in the docProps directory.
  """
  def get_docProps_dir(data) do
    [get_docProps_app_xml(data), get_docProps_core_xml(data)]
  end

  @spec get__rels_dotrels(Workbook.t()) :: zip_tuple
  @doc ~S"""
  Returns the filename '_rels/.rels' and it's content as a tuple
  """
  def get__rels_dotrels(_) do
    {'_rels/.rels', XMLTemplates.rels_dotrels()}
  end

  @spec get__rels_dir(Workbook.t()) :: list(zip_tuple)
  @doc ~S"""
  Returns files in the _rels/ directory.
  """
  def get__rels_dir(data) do
    [get__rels_dotrels(data)]
  end

  @spec get_xl_rels_dir(any, [SheetCompInfo.t()], non_neg_integer) :: list(zip_tuple)
  def get_xl_rels_dir(_, sheetCompInfos, next_rId) do
    [
      {'xl/_rels/workbook.xml.rels',
       ~S"""
       <?xml version="1.0" encoding="UTF-8"?>
       <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
         <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
       """ <>
         XMLTemplates.make_xl_rel_sheets(sheetCompInfos) <>
         """
           <Relationship Id="rId#{next_rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
         </Relationships>
         """}
    ]
  end

  @spec get_xl_styles_xml(WorkbookCompInfo.t()) :: zip_tuple
  def get_xl_styles_xml(wci) do
    {'xl/styles.xml', XMLTemplates.make_xl_styles(wci)}
  end

  @spec get_xl_workbook_xml(Workbook.t(), [SheetCompInfo.t()]) :: zip_tuple
  def get_xl_workbook_xml(data, sheetCompInfos) do
    {'xl/workbook.xml', XMLTemplates.make_workbook_xml(data, sheetCompInfos)}
  end

  @spec get_xl_sharedStrings_xml(any, WorkbookCompInfo.t()) :: zip_tuple
  def get_xl_sharedStrings_xml(_, wci) do
    {'xl/sharedStrings.xml',
     XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(wci.stringdb))}
  end

  @spec sheet_full_path(SheetCompInfo.t()) :: list(char)
  defp sheet_full_path(sci) do
    String.to_charlist("xl/worksheets/#{sci.filename}")
  end

  @spec get_xl_worksheets_dir(Workbook.t(), WorkbookCompInfo.t()) :: list(zip_tuple)
  def get_xl_worksheets_dir(data, wci) do
    sheets = data.sheets

    Enum.zip(sheets, wci.sheet_info)
    |> Enum.map(fn {s, sci} ->
      {sheet_full_path(sci), XMLTemplates.make_sheet(s, wci)}
    end)
  end

  def get_contentTypes_xml(_, wci) do
    {'[Content_Types].xml', XMLTemplates.make_contenttypes_xml(wci)}
  end

  def get_xl_dir(data, wci) do
    sheet_comp_infos = wci.sheet_info
    next_free_xl_rid = wci.next_free_xl_rid

    [
      get_xl_styles_xml(wci),
      get_xl_sharedStrings_xml(data, wci),
      get_xl_workbook_xml(data, sheet_comp_infos)
    ] ++
      get_xl_rels_dir(data, sheet_comp_infos, next_free_xl_rid) ++
      get_xl_worksheets_dir(data, wci)
  end
end
