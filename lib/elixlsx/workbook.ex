defmodule Elixlsx.Workbook do
  alias __MODULE__
  alias Elixlsx.Sheet

  defstruct sheets: [], datetime: nil
  @type t :: %Workbook{
    sheets: nonempty_list(Sheet.t),
    datetime: Elixlsx.Util.calendar
  }

  def append_sheet(workbook, sheet) do
    update_in workbook.sheets, &(&1 ++ [sheet])
  end

  def insert_sheet(workbook, sheet, at \\ 0) do
    update_in workbook.sheets, &(List.insert_at &1, at, sheet)
  end
end
