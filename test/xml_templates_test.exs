defmodule Elixlsx.XMLTemplatesTest do
  use ExUnit.Case

  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Sheet
  alias Elixlsx.XMLTemplates

  describe("make_sheet") do
    setup do
      [sheet: Sheet.with_name("Sheet1"), wci: %WorkbookCompInfo{}]
    end

    test "with default values", %{sheet: sheet, wci: wci} do
      assert XMLTemplates.make_sheet(sheet, wci) =~ "</worksheet>"
    end

    test "with column widths set", %{sheet: sheet, wci: wci} do
      xml = sheet
      |> Sheet.set_col_width("A", 10)
      |> Sheet.set_col_width("B", 12)
      |> XMLTemplates.make_sheet(wci)
      assert xml =~ "<col min=\"1\" max=\"1\" width=\"10\" customWidth=\"1\"/>"
      assert xml =~ "<col min=\"2\" max=\"2\" width=\"12\" customWidth=\"1\"/>"
    end
  end
end
