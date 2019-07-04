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
            filldb: %Compiler.FillDB{},
            cellstyledb: %Compiler.CellStyleDB{},
            numfmtdb: %Compiler.NumFmtDB{},
            borderstyledb: %Compiler.BorderStyleDB{},
            next_free_xl_rid: nil

  @type t :: %Compiler.WorkbookCompInfo{
          sheet_info: [Compiler.SheetCompInfo.t()],
          stringdb: Compiler.StringDB.t(),
          fontdb: Compiler.FontDB.t(),
          filldb: Compiler.FillDB.t(),
          cellstyledb: Compiler.CellStyleDB.t(),
          numfmtdb: Compiler.NumFmtDB.t(),
          borderstyledb: Compiler.BorderStyleDB.t(),
          next_free_xl_rid: non_neg_integer
        }
end
