defmodule Elixlsx.Sheet do
  alias __MODULE__
  alias Elixlsx.Sheet
  alias Elixlsx.Util

  @moduledoc ~S"""
  Describes a single sheet with a given name. The name can be up to 31
  characters long.

  The rows property is a list, each corresponding to a
  row (from the top), of lists, each corresponding to
  a column (from the left), of contents.

  Content may be:

  - a String.t (unicode),
  - a number, or
  - a list [String|number, property_list...]

  The property list describes formatting options for that
  cell. See `Font.from_props/1` for a list of options.
  """
  defstruct name: "",
            rows: [],
            col_widths: %{},
            row_heights: %{},
            group_cols: [],
            group_rows: [],
            merge_cells: [],
            pane_freeze: nil,
            show_grid_lines: true,
            data_validations: []

  @type t :: %Sheet{
          name: String.t(),
          rows: list(list(any())),
          col_widths: %{pos_integer => number},
          row_heights: %{pos_integer => number},
          group_cols: list(rowcol_group),
          group_rows: list(rowcol_group),
          merge_cells: [{String.t(), String.t()}],
          pane_freeze: {number, number} | nil,
          show_grid_lines: boolean(),
          data_validations: list({String.t(), String.t(), list(String.t()) | String.t()})
        }
  @type rowcol_group :: Range.t() | {Range.t(), opts :: keyword}

  @doc ~S"""
  Create a sheet with a sheet name.

  The name can be up to 31 characters long.
  """
  @spec with_name(String.t()) :: Sheet.t()
  def with_name(name) do
    %Sheet{name: name}
  end

  defp split_cell_content_props(cell) do
    cond do
      is_list(cell) ->
        {hd(cell), tl(cell)}

      true ->
        {cell, []}
    end
  end

  @doc ~S"""
  Returns a "CSV" representation of the Sheet.

  This is mainly used for doctests and does not generate valid CSV (yet).
  """
  def to_csv_string(sheet) do
    Enum.map_join(sheet.rows, "\n", fn row ->
      Enum.map_join(row, ",", fn cell ->
        {content, _} = split_cell_content_props(cell)

        case content do
          nil -> ""
          _ -> to_string(content)
        end
      end)
    end)
  end

  @spec set_cell(Sheet.t(), String.t(), any(), Keyword.t()) :: Sheet.t()
  @doc ~S"""
  Set a cell indexed by excel coordinates.

  ## Examples

      iex> %Elixlsx.Sheet{} |>
      ...> Elixlsx.Sheet.set_cell("C1", "Hello World",
      ...>                bold: true, underline: true) |>
      ...> Elixlsx.Sheet.to_csv_string
      ",,Hello World"

  """

  def set_cell(sheet, index, content, opts \\ []) when is_binary(index) do
    {row, col} = Util.from_excel_coords0(index)
    set_at(sheet, row, col, content, opts)
  end

  @spec set_at(Sheet.t(), non_neg_integer, non_neg_integer, any(), Keyword.t()) :: Sheet.t()
  @doc ~S"""
  Set a cell at a given row/column index.

  Indizes start at 0.

  ## Examples

      iex> %Elixlsx.Sheet{} |>
      ...> Elixlsx.Sheet.set_at(0, 2, "Hello World",
      ...>                bold: true, underline: true) |>
      ...> Elixlsx.Sheet.to_csv_string
      ",,Hello World"

  """
  def set_at(sheet, rowidx, colidx, content, opts \\ [])
      when is_number(rowidx) and is_number(colidx) do
    cond do
      length(sheet.rows) <= rowidx ->
        # append new rows, call self again with new sheet
        n_new_rows = rowidx - length(sheet.rows)
        new_rows = 0..n_new_rows |> Enum.map(fn _ -> [] end)

        update_in(sheet.rows, &(&1 ++ new_rows))
        |> set_at(rowidx, colidx, content, opts)

      length(Enum.at(sheet.rows, rowidx)) <= colidx ->
        n_new_cols = colidx - length(Enum.at(sheet.rows, rowidx))
        new_cols = 0..n_new_cols |> Enum.map(fn _ -> nil end)
        new_row = Enum.at(sheet.rows, rowidx) ++ new_cols

        update_in(sheet.rows, &List.replace_at(&1, rowidx, new_row))
        |> set_at(rowidx, colidx, content, opts)

      true ->
        update_in(sheet.rows, fn rows ->
          List.update_at(rows, rowidx, fn cols ->
            List.replace_at(cols, colidx, [content | opts])
          end)
        end)
    end
  end

  @spec set_col_width(Sheet.t(), String.t(), number) :: Sheet.t()
  @doc ~S"""
  Set the column width for a given column.

  Column is indexed by name ("A", ...)
  """
  def set_col_width(sheet, column, width) do
    update_in(
      sheet.col_widths,
      &Map.put(&1, Util.decode_col(column), width)
    )
  end

  @spec set_row_height(Sheet.t(), number, number) :: Sheet.t()
  @doc ~S"""
  Set the row height for a given row.

  Row is indexed starting from 1
  """
  def set_row_height(sheet, row_idx, height) do
    update_in(
      sheet.row_heights,
      &Map.put(&1, row_idx, height)
    )
  end

  @spec group_cols(Sheet.t(), String.t(), String.t()) :: Sheet.t()
  @doc ~S"""
  Group given column range. (i.e. increase outline level by one)

  Column is indexed by name ("A", ...)

  ## Options

    - `collapsed`: if true, collapse this group.
  """
  def group_cols(sheet, first_col, last_col, opts \\ []) do
    col_range = Range.new(Util.decode_col(first_col), Util.decode_col(last_col))
    new_group = if opts === [], do: col_range, else: {col_range, opts}
    update_in(sheet.group_cols, fn groups -> groups ++ [new_group] end)
  end

  @spec group_rows(Sheet.t(), pos_integer, pos_integer) :: Sheet.t()
  @doc ~S"""
  Group given row range. (i.e. increase outline level by one)

  Row is indexed starting from 1.

  ## Options

    - `collapsed`: if true, collapse this group.
  """
  def group_rows(sheet, first_row_idx, last_row_idx, opts \\ []) do
    row_range = Range.new(first_row_idx, last_row_idx)
    new_group = if opts === [], do: row_range, else: {row_range, opts}
    update_in(sheet.group_rows, fn groups -> groups ++ [new_group] end)
  end

  @spec set_pane_freeze(Sheet.t(), number, number) :: Sheet.t()
  @doc ~S"""
  Set the pane freeze at the given row and column.

  Row and column are indexed starting from 1.
  Special value 0 means no freezing, e.g. {1, 0} will freeze first row and no columns.
  """
  def set_pane_freeze(sheet, row_idx, col_idx) do
    %{sheet | pane_freeze: {row_idx, col_idx}}
  end

  @spec remove_pane_freeze(Sheet.t()) :: Sheet.t()
  @doc ~S"""
  Removes any pane freezing that has been set.
  """
  def remove_pane_freeze(sheet) do
    %{sheet | pane_freeze: nil}
  end

  @spec add_data_validations(Sheet.t(), String.t(), String.t(), list(String.t())) :: Sheet.t()
  def add_data_validations(sheet, start_cell, end_cell, values) do
    %{sheet | data_validations: [{start_cell, end_cell, values} | sheet.data_validations]}
  end
end
