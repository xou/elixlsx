defmodule Elixlsx.XMLTemplatesTest do
  use ExUnit.Case

  alias Elixlsx.{Compiler, Compiler.WorkbookCompInfo, Sheet, Workbook, XMLTemplates}

  describe("make_sheet") do
    setup do
      [sheet: Sheet.with_name("Sheet1"), wci: %WorkbookCompInfo{}, workbook: %Workbook{}]
    end

    test "with default values", %{sheet: sheet, wci: wci} do
      assert XMLTemplates.make_sheet(sheet, wci) =~ "</worksheet>"
    end

    test "with column attrs set", %{sheet: sheet, workbook: workbook} do
      sheet = sheet
      |> Sheet.set_col_width("A", 10)
      |> Sheet.set_col("B", bg_color: "#FFFF00", num_format: "mmm-yyyy", width: 12)
      wci = Compiler.make_workbook_comp_info(Workbook.insert_sheet(workbook, sheet))
      xml = XMLTemplates.make_sheet(sheet, wci)
      assert xml =~ "<col min=\"1\" max=\"1\" width=\"10\" customWidth=\"1\"/>"
      assert xml =~ "<col min=\"2\" max=\"2\" width=\"12\" customWidth=\"1\" style=\"1\"/>"
    end
  end
end
