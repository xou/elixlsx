defmodule Elixlsx do
  @moduledoc ~S"""
  Elixlsx is a writer for the MS Excel OpenXML format (`.xlsx`).

  # Quick Overview

  The `write_to/2` function takes a `Elixlsx.Workbook` object
  and a filename. A Workbook is a collection of `Elixlsx.Sheet`s with
  (currently only) a *creation date*.

  See the example.exs file for usage instructions.

  # Hacking / Technical overview

  XLSX stores potentially repeating values in databases, most
  notably `sharedStrings.xml` and `styles.xml`. In these databases,
  each element is assigned a unique ID which is then referenced
  later. IDs are consecutive and correspond to the (0-indexed)
  position in the database (except for number/date formattings,
  where the ID is explicitly given in the attribute and needs to
  be at least 164).

  The sharedStrings database is built up using the
  `Elixlsx.Compiler.StringDB` module. Pre-compilation, all cells
  are *folded* over, producing the StringDB struct which assigns
  each string a unique ID. The StringDB is part of the
  `Elixlsx.Compiler.WorkbookCompInfo` struct, which is passed to
  the XML generating function, which then `get_id`'s the ID
  associated with the string found in the cell.

  For `styles.xml`, the procedure is in general the same, but slightly
  more complicated since elements can reference other elements in
  the same file. The `Elixlsx.Style.CellStyle` element is the
  combination of sub-styles (`Elixlsx.Style.Font`, `Elixlsx.Style.NumFmt`,
  ...). A call to register_all creates the (unique) entries in the
  sub-style databases (`Elixlsx.Compiler.FontDB`, `Elixlsx.Compiler.NumFmtDB`).
  Afterwards, each unique combination of substyles is assigned an ID
  in `Elixlsx.Compiler.CellStyleDB`. During XML generation, the &lt;xf&gt;
  elements reference the individual sub-style IDs, and the actual cell
  element references the &lt;xf&gt; id.
  """

  @doc ~S"""
  Write a Workbook object to the given filename.
  """
  @spec write_to(Elixlsx.Workbook.t(), String.t()) :: {:ok, charlist} | {:error, any()}
  def write_to(workbook, filename) do
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    case :zip.create(to_charlist(filename), Elixlsx.Writer.create_files(workbook, wci)) do
      {:ok, _} -> {:ok, filename}
      {:error, error} -> {:error, error}
    end
  end

  @doc ~S"""
  Write a Workbook object to the binary.

  Returns a tuple containing a filename and the binary
  """
  @spec write_to_memory(Elixlsx.Workbook.t(), String.t()) ::
          {:ok, {charlist, binary}} | {:error, any()}
  def write_to_memory(workbook, filename) do
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    case :zip.create(to_charlist(filename), Elixlsx.Writer.create_files(workbook, wci), [:memory]) do
      {:ok, _} -> {:ok, filename}
      {:error, error} -> {:error, error}
    end
  end
end
