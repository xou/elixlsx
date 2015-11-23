defmodule Workbook do
  defstruct sheets: [], datetime: nil
  @type t :: %Workbook{
      sheets: nonempty_list(Sheet.t),
      datetime: Elixlsx.Util.calendar
  }
end

defmodule Sheet do
  @moduledoc ~S"""
  Describes a single sheet with a given name.
  The rows property is a list, each corresponding to a
  row (from the top), of lists, each corresponding to
  a column (from the left), of contents.

  Content may be
  - a String.t (unicode),
  - a number, or
  - a list [String|number, property_list...]

  The property list describes formatting options for that
  cell. See Font.from_props/1 for a list of options.
  """
  defstruct name: "", rows: [], sheetCompInfo: nil
  @type t :: %Sheet {
    name: String.t,
    rows: list(list(any())),
  }
end

defmodule Elixlsx do
  def write_to(workbook, filename) do
    wci = Elixlsx.Compiler.make_workbook_comp_info workbook
    :zip.create(filename, Elixlsx.Writer.create_files(workbook, wci))
  end
end

