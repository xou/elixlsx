defmodule Elixlsx do
  @doc ~S"""
  Write a Workbook object to the given filename
  """
  @spec write_to(Elixlsx.Workbook.t, String.t) :: {:ok, String.t} | {:error, any()}
  def write_to(workbook, filename) do
    wci = Elixlsx.Compiler.make_workbook_comp_info workbook
    :zip.create(filename, Elixlsx.Writer.create_files(workbook, wci))
  end
end

