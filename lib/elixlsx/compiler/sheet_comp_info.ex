defmodule Elixlsx.Compiler.SheetCompInfo do
  alias Elixlsx.Compiler.SheetCompInfo

  @moduledoc ~S"""
  Compilation info for a sheet, to be filled during the actual
  write process.
  """
  defstruct rId: "", filename: "sheet1.xml", sheetId: 0

  @type t :: %SheetCompInfo{
          rId: String.t(),
          filename: String.t(),
          sheetId: non_neg_integer
        }

  @spec make(non_neg_integer, non_neg_integer) :: SheetCompInfo.t()
  def make(sheetidx, rId) do
    %SheetCompInfo{
      rId: "rId" <> to_string(rId),
      filename: "sheet" <> to_string(sheetidx) <> ".xml",
      sheetId: sheetidx
    }
  end
end
