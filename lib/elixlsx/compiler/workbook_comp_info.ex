defmodule Elixlsx.Compiler.WorkbookCompInfo do
  alias Elixlsx.Compiler

  @moduledoc ~S"""
  This module aggregates information about the metainformation
  required to generate the XML file.

  It is used as the aggregator when folding over the individual
  cells and images.
  """
  defstruct sheet_info: nil,
            drawing_info: nil,
            stringdb: %Compiler.StringDB{},
            fontdb: %Compiler.FontDB{},
            filldb: %Compiler.FillDB{},
            cellstyledb: %Compiler.CellStyleDB{},
            numfmtdb: %Compiler.NumFmtDB{},
            borderstyledb: %Compiler.BorderStyleDB{},
            drawingdb: %Compiler.DrawingDB{},
            next_free_xl_rid: nil

  @type t :: %Compiler.WorkbookCompInfo{
          sheet_info: [Compiler.SheetCompInfo.t()],
          drawing_info: [Compiler.DrawingCompInfo.t()],
          stringdb: Compiler.StringDB.t(),
          fontdb: Compiler.FontDB.t(),
          filldb: Compiler.FillDB.t(),
          cellstyledb: Compiler.CellStyleDB.t(),
          numfmtdb: Compiler.NumFmtDB.t(),
          borderstyledb: Compiler.BorderStyleDB.t(),
          drawingdb: Compiler.DrawingDB.t(),
          next_free_xl_rid: non_neg_integer
        }
end
