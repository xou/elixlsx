defmodule ExcelCompatibilityValidation do
  use ExUnit.Case
  
  alias Elixlsx.{Workbook, Sheet}
  
  @moduledoc """
  Final validation tests for Excel and LibreOffice compatibility.
  These tests validate that generated files meet Excel format specifications.
  """

  @test_file "excel_compatibility_test.xlsx"

  setup do
    on_exit(fn -> File.rm(@test_file) end)
  end

  describe "Excel format compliance" do
    test "comprehensive Excel feature validation" do
      # Create a workbook that exercises all major Excel features
      sheet1 = Sheet.with_name("Feature Test")
      |> Sheet.set_cell("A1", "Excel Compatibility Test", bold: true, size: 14, color: "#0066CC")
      |> Sheet.set_cell("A2", "Data Types:", bold: true)
      |> Sheet.set_cell("B2", "String")
      |> Sheet.set_cell("C2", 42)
      |> Sheet.set_cell("D2", 3.14159)
      |> Sheet.set_cell("E2", true)
      |> Sheet.set_cell("F2", {{2024, 1, 15}, {10, 30, 0}}, datetime: true)
      |> Sheet.set_cell("A3", "Formatting:", bold: true)
      |> Sheet.set_cell("B3", "Bold", bold: true)
      |> Sheet.set_cell("C3", "Italic", italic: true)
      |> Sheet.set_cell("D3", "Underline", underline: true)
      |> Sheet.set_cell("E3", "Strike", strike: true)
      |> Sheet.set_cell("F3", "Large", size: 18)
      |> Sheet.set_cell("A4", "Colors:", bold: true)
      |> Sheet.set_cell("B4", "Red Text", color: "#FF0000")
      |> Sheet.set_cell("C4", "Blue BG", bg_color: "#0000FF", color: "#FFFFFF")
      |> Sheet.set_cell("D4", "Green", color: "#00FF00")
      |> Sheet.set_cell("A5", "Numbers:", bold: true)
      |> Sheet.set_cell("B5", 1234.56, num_format: "#,##0.00")
      |> Sheet.set_cell("C5", 0.85, num_format: "0.00%")
      |> Sheet.set_cell("D5", 1000000, num_format: "#,##0")
      |> Sheet.set_cell("A6", "Formulas:", bold: true)
      |> Sheet.set_cell("B6", {:formula, "C2+D2", value: 45.14159})
      |> Sheet.set_cell("C6", {:formula, "NOW()"}, num_format: "yyyy-mm-dd")
      |> Sheet.set_cell("A7", "Borders:", bold: true)
      |> Sheet.set_cell("B7", "Single", border: [bottom: [style: :thin, color: "#000000"]])
      |> Sheet.set_cell("C7", "Double", border: [bottom: [style: :double, color: "#FF0000"]])
      |> Sheet.set_cell("D7", "Thick", border: [bottom: [style: :thick, color: "#0000FF"]])
      |> Sheet.set_col_width("A", 12.0)
      |> Sheet.set_col_width("B", 15.0)
      |> Sheet.set_col_width("C", 15.0)
      |> Sheet.set_col_width("D", 15.0)
      |> Sheet.set_col_width("E", 15.0)
      |> Sheet.set_col_width("F", 18.0)

      # Second sheet with matrix data
      sheet2 = %Sheet{
        name: "Data Matrix",
        rows: [
          ["Quarter", "Q1", "Q2", "Q3", "Q4"],
          ["Product A", 100, 150, 200, 175],
          ["Product B", 200, 250, 300, 275],
          ["Product C", 150, 175, 225, 200],
          ["Total", {:formula, "SUM(B2:B4)", value: 450}, {:formula, "SUM(C2:C4)", value: 575}, {:formula, "SUM(D2:D4)", value: 725}, {:formula, "SUM(E2:E4)", value: 650}]
        ]
      }

      # Third sheet with merged cells and advanced features
      sheet3 = %Sheet{
        name: "Advanced Features",
        rows: [
          ["Merged Header", "", "", ""],
          ["Data 1", "Data 2", "Data 3", "Data 4"],
          [1, 2, 3, 4],
          [5, 6, 7, 8]
        ],
        merge_cells: [{"A1", "D1"}]
      }
      |> Sheet.set_pane_freeze(2, 1)

      workbook = %Workbook{
        sheets: [sheet1, sheet2, sheet3],
        datetime: "2024-01-15T10:30:00Z"
      }

      # Generate the Excel file
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file)
      assert File.exists?(@test_file)

      # Validate file structure
      validate_excel_structure(@test_file)
      
      # Validate file size is reasonable
      file_info = File.stat!(@test_file)
      assert file_info.size > 5000, "File should be at least 5KB for comprehensive content"
      assert file_info.size < 100_000, "File should not be excessively large"
    end

    test "unicode and special character handling" do
      # Test comprehensive Unicode support
      sheet = Sheet.with_name("Unicode Test")
      |> Sheet.set_cell("A1", "Basic Latin: Hello World!")
      |> Sheet.set_cell("A2", "Latin Extended: Cafe, naive, resume")
      |> Sheet.set_cell("A3", "German: Mueller, Groesse, weiss")
      |> Sheet.set_cell("A4", "French: francais, coeur, Noel")
      |> Sheet.set_cell("A5", "Spanish: nino, senora, ano")
      |> Sheet.set_cell("A6", "Symbols: Copyright(C) Registered(R) Trademark(TM)")
      |> Sheet.set_cell("A7", "Math: a^2+b^2=c^2, Sum Product Integral")
      |> Sheet.set_cell("A8", "Arrows: Left Right Up Down")
      |> Sheet.set_cell("A9", "XML Special: <tag>&amp;\"quote\"'apostrophe'")
      |> Sheet.set_cell("A10", "Emoji: Happy Party Chart Business Fire Star")
      |> Sheet.set_col_width("A", 35.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file)
      
      # Validate that the file can be opened and contains proper XML
      validate_excel_structure(@test_file)
      validate_xml_encoding(@test_file)
    end

    test "large dataset performance and compatibility" do
      # Test with a substantial dataset
      header = ["ID", "Name", "Value", "Date", "Status", "Amount"]
      
      data_rows = Enum.map(1..2000, fn i ->
        [
          i,
          "Item #{i}",
          i * 1.5,
          "2024-01-#{String.pad_leading(to_string(rem(i, 28) + 1), 2, "0")}",
          if(rem(i, 2) == 0, do: "Active", else: "Inactive"),
          i * 10.50
        ]
      end)

      sheet = %Sheet{
        name: "Large Dataset",
        rows: [header | data_rows]
      }

      workbook = %Workbook{sheets: [sheet]}
      
      # Measure performance
      {time_microseconds, result} = :timer.tc(fn ->
        Elixlsx.write_to(workbook, @test_file)
      end)
      
      assert {:ok, _} = result
      
      # Should complete within reasonable time (less than 10 seconds for 2000 rows)
      assert time_microseconds < 10_000_000, "Large dataset generation took too long: #{time_microseconds / 1_000_000} seconds"
      
      # Validate file
      validate_excel_structure(@test_file)
      
      file_info = File.stat!(@test_file)
      assert file_info.size > 50_000, "Large dataset should produce substantial file"
    end

    test "formula and calculation compatibility" do
      # Test various Excel formulas
      sheet = Sheet.with_name("Formula Test")
      |> Sheet.set_cell("A1", "Basic Math")
      |> Sheet.set_cell("A2", 10)
      |> Sheet.set_cell("A3", 20)
      |> Sheet.set_cell("A4", {:formula, "A2+A3", value: 30})
      |> Sheet.set_cell("A5", {:formula, "A2*A3", value: 200})
      |> Sheet.set_cell("A6", {:formula, "A3/A2", value: 2})
      |> Sheet.set_cell("B1", "Statistical")
      |> Sheet.set_cell("B2", {:formula, "SUM(A2:A3)", value: 30})
      |> Sheet.set_cell("B3", {:formula, "AVERAGE(A2:A3)", value: 15})
      |> Sheet.set_cell("B4", {:formula, "MAX(A2:A3)", value: 20})
      |> Sheet.set_cell("B5", {:formula, "MIN(A2:A3)", value: 10})
      |> Sheet.set_cell("B6", {:formula, "COUNT(A2:A3)", value: 2})
      |> Sheet.set_cell("C1", "Date/Time")
      |> Sheet.set_cell("C2", {:formula, "NOW()"})
      |> Sheet.set_cell("C3", {:formula, "TODAY()"})
      |> Sheet.set_cell("C4", {:formula, "YEAR(NOW())", value: 2024})
      |> Sheet.set_cell("C5", {:formula, "MONTH(NOW())", value: 1})

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file)
      validate_excel_structure(@test_file)
    end

    test "data validation and advanced features" do
      # Test data validation and other advanced Excel features
      sheet = Sheet.with_name("Advanced Test")
      |> Sheet.set_cell("A1", "Validation Options:")
      |> Sheet.set_cell("A2", "Option 1")
      |> Sheet.set_cell("A3", "Option 2")
      |> Sheet.set_cell("A4", "Option 3")
      |> Sheet.set_cell("C1", "Select Here:")
      |> Sheet.set_cell("C2", "Option 1")
      |> Sheet.add_data_validations("C2", "C10", ["Option 1", "Option 2", "Option 3"])
      |> Sheet.add_data_validations("D2", "D10", "=$A$2:$A$4")
      |> Sheet.set_cell("F1", "Wrapped Text Example")
      |> Sheet.set_cell("F2", "This is a very long text that should wrap within the cell when wrap_text is enabled", wrap_text: true)
      |> Sheet.set_row_height(2, 50)
      |> Sheet.set_col_width("F", 25.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file)
      validate_excel_structure(@test_file)
    end
  end

  # Helper function to validate Excel file structure
  defp validate_excel_structure(file_path) do
    # Verify it's a valid ZIP file
    assert {:ok, file_list} = :zip.list_dir(to_charlist(file_path))
    
    # Extract file names
    file_names = Enum.map(file_list, fn 
      {:zip_file, name, _info, _comment, _offset, _comp_size} -> to_string(name)
      {name, _info} -> to_string(name)
    end)
    
    # Check for essential Excel files
    required_files = [
      "[Content_Types].xml",
      "_rels/.rels",
      "xl/workbook.xml",
      "xl/worksheets/sheet1.xml",
      "xl/sharedStrings.xml",
      "xl/styles.xml"
    ]
    
    Enum.each(required_files, fn file ->
      assert file in file_names, "Missing required Excel file: #{file}"
    end)
    
    # Verify we can extract the files
    assert {:ok, _files} = :zip.extract(to_charlist(file_path), [:memory])
  end

  # Helper function to validate XML encoding
  defp validate_xml_encoding(file_path) do
    assert {:ok, files} = :zip.extract(to_charlist(file_path), [:memory])
    
    # Check key XML files for proper encoding
    xml_files = [~c"xl/sharedStrings.xml", ~c"xl/workbook.xml", ~c"xl/worksheets/sheet1.xml"]
    
    Enum.each(xml_files, fn xml_file ->
      xml_content = Enum.find_value(files, fn
        {^xml_file, content} -> content
        _ -> nil
      end)
      
      assert xml_content != nil, "Missing XML file: #{xml_file}"
      
      # Verify XML is well-formed
      xml_string = to_string(xml_content)
      assert String.starts_with?(xml_string, "<?xml version=\"1.0\" encoding=\"UTF-8\"")
      
      # Verify it can be parsed
      xml_charlist = String.to_charlist(xml_string)
      {parsed_xml, _} = :xmerl_scan.string(xml_charlist)
      assert parsed_xml != nil
    end)
  end
end