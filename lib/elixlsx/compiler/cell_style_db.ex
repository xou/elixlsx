defmodule Elixlsx.Compiler.CellStyleDB do
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.FontDB
  alias Elixlsx.Compiler.FillDB
  alias Elixlsx.Compiler.NumFmtDB
  alias Elixlsx.Compiler.BorderStyleDB
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Compiler.DBUtil

  defstruct cellstyles: %{}, element_count: 0

  @type t :: %CellStyleDB {
    cellstyles: %{Elixlsx.Style.CellStyle.t => non_neg_integer},
    element_count: non_neg_integer
  }


  def register_style(cellstyledb, style) do
    case Map.fetch(cellstyledb.cellstyles, style) do
      :error ->
        # add +1 here already since "0" refers to the default style
        csdb = update_in cellstyledb.cellstyles,
                  &(Map.put &1, style, (cellstyledb.element_count + 1))
        update_in csdb.element_count, &(&1 + 1)
      {:ok, _} ->
        cellstyledb
    end
  end

  def get_id(cellstyledb, style) do
    case Map.fetch(cellstyledb.cellstyles, style) do
      :error ->
        raise %ArgumentError{message: "Could not find key in styledb: " <> inspect(style)}
      {:ok, key} ->
        key
    end
  end

  def id_sorted_styles(cellstyledb), do: DBUtil.id_sorted_values(cellstyledb.cellstyles)

  @doc ~S"""
  Recursively register all fonts, numberformat,
  border* and fill* properties (*=TBD)
  in the WorkbookCompInfo structure.
  """
  @spec register_all(WorkbookCompInfo.t) :: WorkbookCompInfo.t
  def register_all(wci) do
    Enum.reduce wci.cellstyledb.cellstyles, wci, fn ({style, _}, wci) ->
      wci = if is_nil(style.font) do
        wci
      else
        update_in(wci.fontdb, &(FontDB.register_font &1, style.font))
      end
      wci = if is_nil(style.fill) do
        wci
      else
        update_in(wci.filldb, &(FillDB.register_fill &1, style.fill))
      end
      wci = if is_nil(style.numfmt) do
        wci
      else
        update_in(wci.numfmtdb, &(NumFmtDB.register_numfmt &1, style.numfmt))
      end
      wci = if is_nil(style.border) do
        wci
      else
        update_in(wci.borderstyledb, &(BorderStyleDB.register_border &1, style.border))
      end

      wci
      # TODO: update_in wci.borderstyledb ...; wci.fillstyledb...
    end
  end
end
