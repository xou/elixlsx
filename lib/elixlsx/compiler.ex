defmodule Elixlsx.Compiler do
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Compiler.SheetCompInfo
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.LinkDB
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.XML
  alias Elixlsx.Sheet

  @doc ~S"""
  Accepts a list of Sheets and the next free relationship ID.

  Returns a tuple containing a list of SheetCompInfo's and the next free
  relationship ID.
  """
  @spec make_sheet_info(nonempty_list(Sheet.t()), non_neg_integer) ::
          {list(SheetCompInfo.t()), non_neg_integer}
  def make_sheet_info(sheets, init_rId) do
    {sheetCompInfos, _, nextrID} =
      List.foldl(sheets, {[], 1, init_rId}, fn sheet, {sci, idx, rId} ->
        sheetcomp = SheetCompInfo.make(idx, rId)

        {comp, rId} = complink_from_rows(sheetcomp, sheet.rows, rId + 1)

        {[comp | sci], idx + 1, rId}
      end)

    {Enum.reverse(sheetCompInfos), nextrID}
  end

  defp complink_from_rows(sci, rows, rid) do
    List.foldl(rows, {sci, rid}, fn cols, {sci, rid} ->
      List.foldl(cols, {sci, rid}, fn cell, {sci, rid} ->
        complink_cell_pass(sci, rid, cell)
      end)
    end)
  end

  defp complink_cell_pass(sci, rid, [{:link, {url, _}} | _]) do
    {update_in(sci.linkdb, &LinkDB.register_link(&1, url, rid)), rid + 1}
  end

  defp complink_cell_pass(%SheetCompInfo{} = sci, rid, _cell) do
    {sci, rid}
  end

  def compinfo_cell_pass_value(wci, {:link, {_url, value}}) do
    compinfo_cell_pass_value(wci, value)
  end

  def compinfo_cell_pass_value(wci, value) do
    cond do
      is_binary(value) && XML.valid?(value) ->
        update_in(wci.stringdb, &StringDB.register_string(&1, value))

      true ->
        wci
    end
  end

  def compinfo_cell_pass_style(wci, props) do
    update_in(
      wci.cellstyledb,
      &CellStyleDB.register_style(
        &1,
        Elixlsx.Style.CellStyle.from_props(props)
      )
    )
  end

  @spec compinfo_cell_pass(WorkbookCompInfo.t(), any) :: WorkbookCompInfo.t()
  def compinfo_cell_pass(wci, cell) do
    cond do
      is_list(cell) ->
        wci
        |> compinfo_cell_pass_value(hd(cell))
        |> compinfo_cell_pass_style(tl(cell))

      true ->
        # no style information attached in this cell
        compinfo_cell_pass_value(wci, cell)
    end
  end

  @spec compinfo_from_rows(WorkbookCompInfo.t(), list(list(any()))) :: WorkbookCompInfo.t()
  def compinfo_from_rows(wci, rows) do
    List.foldl(rows, wci, fn cols, wci ->
      List.foldl(cols, wci, fn cell, wci ->
        compinfo_cell_pass(wci, cell)
      end)
    end)
  end

  @spec compinfo_from_sheets(WorkbookCompInfo.t(), list(Sheet.t())) :: WorkbookCompInfo.t()
  def compinfo_from_sheets(wci, sheets) do
    List.foldl(sheets, wci, fn sheet, wci ->
      compinfo_from_rows(wci, sheet.rows)
    end)
  end

  @first_free_rid 2
  def make_workbook_comp_info(workbook) do
    {sci, next_rId} = make_sheet_info(workbook.sheets, @first_free_rid)

    %WorkbookCompInfo{
      sheet_info: sci,
      next_free_xl_rid: next_rId
    }
    |> compinfo_from_sheets(workbook.sheets)
    |> CellStyleDB.register_all()
  end
end
