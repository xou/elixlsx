defmodule Elixlsx.Compiler do
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Compiler.SheetCompInfo
  alias Elixlsx.Compiler.CellStyleDB
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
    # fold helper. aggregator holds {list(SheetCompInfo), sheetidx, rId}.
    add_sheet = fn _, {sci, idx, rId} ->
      {[SheetCompInfo.make(idx, rId) | sci], idx + 1, rId + 1}
    end

    # TODO probably better to use a zip [1..] |> map instead of fold[l|r]/reverse
    {sheetCompInfos, _, nextrID} = List.foldl(sheets, {[], 1, init_rId}, add_sheet)
    {Enum.reverse(sheetCompInfos), nextrID}
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
