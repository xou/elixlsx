defmodule Elixlsx.Workbook do
  @moduledoc ~S"""
  Root structure for excel files.

  Must contain at least one Elixlsx.Sheet object.

  The datetime property can optionally be set to override
  the "created at" date. It defaults to the current time.

  The defined_names property is a map of defined names in the workbook.
  """
  alias Elixlsx.Sheet
  alias Elixlsx.Workbook

  defstruct sheets: [], datetime: nil, defined_names: %{}

  @type t :: %Workbook{
          sheets: nonempty_list(Sheet.t()),
          datetime: String.t() | integer | nil,
          defined_names: %{String.t() => String.t()}
        }

  @doc "Append a sheet at the end."
  @spec append_sheet(Workbook.t(), Sheet.t()) :: Workbook.t()
  def append_sheet(workbook, sheet) do
    update_in(workbook.sheets, &(&1 ++ [sheet]))
  end

  @doc """
  Insert a sheet at a given position, starting at 0.
  """
  @spec insert_sheet(Workbook.t(), Sheet.t(), non_neg_integer) :: Workbook.t()
  def insert_sheet(workbook, sheet, at \\ 0) do
    update_in(workbook.sheets, &List.insert_at(&1, at, sheet))
  end

  @doc """
  Add a defined name to the workbook.
  """
  @spec add_defined_name(Workbook.t(), String.t(), String.t()) :: Workbook.t()
  def add_defined_name(workbook, name, definition) do
    update_in(workbook.defined_names, &Map.put(&1, name, definition))
  end
end
