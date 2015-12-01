defmodule Elixlsx.Compiler.WorkbookCompInfo do
  alias Elixlsx.Compiler
  @moduledoc ~S"""
  This module aggregates information about the metainformation
  required to generate the XML file.

  It is used as the aggregator when folding over the individual
  cells.
  """
  defstruct sheet_info: nil,
  stringdb: %Compiler.StringDB{},
  fontdb: %Compiler.FontDB{},
  cellstyledb: %Compiler.CellStyleDB{},
  numfmtdb: %Compiler.NumFmtDB{},
  next_free_xl_rid: nil

  @type t :: %Compiler.WorkbookCompInfo{
    sheet_info: Compiler.SheetCompInfo.t,
    stringdb: Compiler.StringDB.t,
    fontdb: Compiler.FontDB.t,
    cellstyledb: Compiler.CellStyleDB.t,
    numfmtdb: Compiler.NumFmtDB.t,
    next_free_xl_rid: non_neg_integer
  }
end
