defmodule IntegrationTest do
  use ExUnit.Case
  
  alias Elixlsx.{Workbook, Sheet}
  
  @moduledoc """
  Comprehensive integration tests for real-world usage scenarios.
  Tests Excel file compatibility and all supported features.
  """

  @test_file_path "test_integration_output.xlsx"
  @memory_test_filename "memory_test.xlsx"

  # Clean up test files after each test
  setup do
    on_exit(fn ->
      File.rm(@test_file_path)
    end)
  end

  describe "Real-world usage scenarios" do
    test "basic data entry and formatting" do
      # Test basic data entry with various data types and formatting
      sheet = Sheet.with_name("Basic Data")
      |> Sheet.set_cell("A1", "Product Name", bold: true, bg_color: "#E0E0E0")
      |> Sheet.set_cell("B1", "Price", bold: true, bg_color: "#E0E0E0")
      |> Sheet.set_cell("C1", "Quantity", bold: true, bg_color: "#E0E0E0")
      |> Sheet.set_cell("D1", "Total", bold: true, bg_color: "#E0E0E0")
      |> Sheet.set_cell("A2", "Widget A")
      |> Sheet.set_cell("B2", 19.99, num_format: "$0.00")
      |> Sheet.set_cell("C2", 5)
      |> Sheet.set_cell("D2", {:formula, "B2*C2", value: 99.95}, num_format: "$0.00")
      |> Sheet.set_cell("A3", "Widget B")
      |> Sheet.set_cell("B3", 29.99, num_format: "$0.00")
      |> Sheet.set_cell("C3", 3)
      |> Sheet.set_cell("D3", {:formula, "B3*C3", value: 89.97}, num_format: "$0.00")
      |> Sheet.set_col_width("A", 15.0)
      |> Sheet.set_col_width("B", 10.0)
      |> Sheet.set_col_width("C", 10.0)
      |> Sheet.set_col_width("D", 12.0)

      workbook = %Workbook{sheets: [sheet]}
      
      # Test file creation
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
      
      # Verify file is not empty and has reasonable size
      file_info = File.stat!(@test_file_path)
      assert file_info.size > 1000  # Should be at least 1KB for a basic Excel file
    end

    test "complex formatting and styling" do
      # Test comprehensive formatting options
      sheet = Sheet.with_name("Formatted Data")
      |> Sheet.set_cell("A1", "Bold Text", bold: true)
      |> Sheet.set_cell("A2", "Italic Text", italic: true)
      |> Sheet.set_cell("A3", "Underlined Text", underline: true)
      |> Sheet.set_cell("A4", "Strike Through", strike: true)
      |> Sheet.set_cell("A5", "Large Font", size: 18)
      |> Sheet.set_cell("A6", "Custom Font", font: "Courier New")
      |> Sheet.set_cell("B1", "Red Text", color: "#FF0000")
      |> Sheet.set_cell("B2", "Blue Background", bg_color: "#0000FF", color: "#FFFFFF")
      |> Sheet.set_cell("B3", "Green Text", color: "#00FF00")
      |> Sheet.set_cell("C1", "Left Aligned", align_horizontal: :left)
      |> Sheet.set_cell("C2", "Center Aligned", align_horizontal: :center)
      |> Sheet.set_cell("C3", "Right Aligned", align_horizontal: :right)
      |> Sheet.set_cell("D1", "Top Aligned", align_vertical: :top)
      |> Sheet.set_cell("D2", "Middle Aligned", align_vertical: :center)
      |> Sheet.set_cell("D3", "Bottom Aligned", align_vertical: :bottom)
      |> Sheet.set_cell("E1", "Wrapped Text that should wrap to multiple lines", wrap_text: true)
      |> Sheet.set_row_height(1, 40)
      |> Sheet.set_col_width("E", 20.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
      
      file_info = File.stat!(@test_file_path)
      assert file_info.size > 2000  # More complex formatting should result in larger file
    end

    test "date and number formatting" do
      # Test various date and number formats
      current_time = {{2024, 1, 15}, {14, 30, 0}}
      unix_timestamp = 1705329000
      
      sheet = Sheet.with_name("Dates and Numbers")
      |> Sheet.set_cell("A1", "Date Formats", bold: true)
      |> Sheet.set_cell("A2", current_time, datetime: true)
      |> Sheet.set_cell("A3", unix_timestamp, datetime: true)
      |> Sheet.set_cell("A4", unix_timestamp, yyyymmdd: true)
      |> Sheet.set_cell("A5", unix_timestamp, yyyymm: true)
      |> Sheet.set_cell("B1", "Number Formats", bold: true)
      |> Sheet.set_cell("B2", 1234.567, num_format: "0.00")
      |> Sheet.set_cell("B3", 1234.567, num_format: "#,##0.00")
      |> Sheet.set_cell("B4", 0.85, num_format: "0.00%")
      |> Sheet.set_cell("B5", 1234567, num_format: "#,##0")
      |> Sheet.set_col_width("A", 20.0)
      |> Sheet.set_col_width("B", 15.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "formulas and calculations" do
      # Test formula functionality
      sheet = Sheet.with_name("Formulas")
      |> Sheet.set_cell("A1", "Value 1")
      |> Sheet.set_cell("A2", "Value 2")
      |> Sheet.set_cell("A3", "Sum")
      |> Sheet.set_cell("A4", "Average")
      |> Sheet.set_cell("A5", "Max")
      |> Sheet.set_cell("A6", "Current Time")
      |> Sheet.set_cell("B1", 10)
      |> Sheet.set_cell("B2", 20)
      |> Sheet.set_cell("B3", {:formula, "B1+B2", value: 30})
      |> Sheet.set_cell("B4", {:formula, "AVERAGE(B1:B2)", value: 15})
      |> Sheet.set_cell("B5", {:formula, "MAX(B1:B2)", value: 20})
      |> Sheet.set_cell("B6", {:formula, "NOW()"}, num_format: "yyyy-mm-dd hh:MM:ss")
      |> Sheet.set_col_width("A", 12.0)
      |> Sheet.set_col_width("B", 15.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "borders and advanced styling" do
      # Test border functionality and advanced styling
      sheet = Sheet.with_name("Borders")
      |> Sheet.set_cell("A1", "Single Border", border: [bottom: [style: :thin, color: "#000000"]])
      |> Sheet.set_cell("A2", "Double Border", border: [bottom: [style: :double, color: "#FF0000"]])
      |> Sheet.set_cell("A3", "Thick Border", border: [bottom: [style: :thick, color: "#0000FF"]])
      |> Sheet.set_cell("B1", "All Borders", 
          border: [
            top: [style: :thin, color: "#000000"],
            bottom: [style: :thin, color: "#000000"],
            left: [style: :thin, color: "#000000"],
            right: [style: :thin, color: "#000000"]
          ])
      |> Sheet.set_cell("C1", "Mixed Borders",
          border: [
            top: [style: :double, color: "#FF0000"],
            bottom: [style: :thick, color: "#00FF00"],
            left: [style: :thin, color: "#0000FF"],
            right: [style: :medium, color: "#FF00FF"]
          ])
      |> Sheet.set_col_width("A", 15.0)
      |> Sheet.set_col_width("B", 15.0)
      |> Sheet.set_col_width("C", 15.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "multiple sheets with different features" do
      # Test multiple sheets with various features
      sheet1 = Sheet.with_name("Summary")
      |> Sheet.set_cell("A1", "Summary Report", bold: true, size: 16)
      |> Sheet.set_cell("A3", "Total Sales:", bold: true)
      |> Sheet.set_cell("B3", {:formula, "Data!D10", value: 1000}, num_format: "$#,##0.00")

      sheet2 = Sheet.with_name("Data")
      |> Sheet.set_cell("A1", "Product", bold: true)
      |> Sheet.set_cell("B1", "Q1", bold: true)
      |> Sheet.set_cell("C1", "Q2", bold: true)
      |> Sheet.set_cell("D1", "Total", bold: true)
      |> Sheet.set_cell("A2", "Product A")
      |> Sheet.set_cell("B2", 250, num_format: "$0.00")
      |> Sheet.set_cell("C2", 300, num_format: "$0.00")
      |> Sheet.set_cell("D2", {:formula, "B2+C2", value: 550}, num_format: "$0.00")
      |> Sheet.set_cell("A10", "Total")
      |> Sheet.set_cell("D10", {:formula, "SUM(D2:D9)", value: 1000}, num_format: "$0.00")

      sheet3 = %Sheet{
        name: "Matrix Data",
        rows: [
          ["Item", "Jan", "Feb", "Mar"],
          ["A", 100, 150, 200],
          ["B", 200, 250, 300],
          ["C", 150, 175, 225]
        ]
      }

      workbook = %Workbook{sheets: [sheet1, sheet2, sheet3]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
      
      # Verify file size is appropriate for multiple sheets
      file_info = File.stat!(@test_file_path)
      assert file_info.size > 3000
    end

    test "data validation functionality" do
      # Test data validation features
      sheet = Sheet.with_name("Validation")
      |> Sheet.set_cell("A1", "Valid Options:")
      |> Sheet.set_cell("A2", "Option A")
      |> Sheet.set_cell("A3", "Option B") 
      |> Sheet.set_cell("A4", "Option C")
      |> Sheet.set_cell("C1", "Select from list:")
      |> Sheet.set_cell("C2", "Option A")
      |> Sheet.add_data_validations("C2", "C10", ["Option A", "Option B", "Option C"])
      |> Sheet.add_data_validations("D2", "D10", "=$A$2:$A$4")

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "merged cells and advanced layout" do
      # Test merged cells and layout features
      sheet = %Sheet{
        name: "Layout",
        rows: [
          ["Merged Header", "", "", ""],
          ["A", "B", "C", "D"],
          [1, 2, 3, 4],
          [5, 6, 7, 8]
        ],
        merge_cells: [{"A1", "D1"}, {"A3", "B4"}]
      }
      |> Sheet.set_pane_freeze(2, 1)  # Freeze first 2 rows and 1 column

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "row and column grouping" do
      # Test grouping functionality
      sheet = %Sheet{
        name: "Grouped Data",
        rows: Enum.map(1..20, fn i -> ["Row #{i}", i * 10, i * 20, i * 30] end),
        group_rows: [{2..5, collapsed: false}, {7..10, collapsed: true}],
        group_cols: [2..4]
      }
      |> Sheet.group_cols("B", "C")

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "unicode and special characters" do
      # Test Unicode support and special characters
      sheet = Sheet.with_name("Unicode Test")
      |> Sheet.set_cell("A1", "English: Hello World")
      |> Sheet.set_cell("A2", "German: Müller & Söhne")
      |> Sheet.set_cell("A3", "French: Café & Résumé")
      |> Sheet.set_cell("A4", "Spanish: Niño & Señora")
      |> Sheet.set_cell("A5", "Symbols: ©®™€£¥")
      |> Sheet.set_cell("A6", "Math: α²+β²=γ²")
      |> Sheet.set_cell("A7", "Emoji: 😀🎉📊")
      |> Sheet.set_cell("B1", "XML Special: <>&\"'")
      |> Sheet.set_cell("B2", "Quotes: \"Hello\" & 'World'")
      |> Sheet.set_col_width("A", 25.0)
      |> Sheet.set_col_width("B", 20.0)

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "memory-based file generation" do
      # Test write_to_memory functionality
      sheet = Sheet.with_name("Memory Test")
      |> Sheet.set_cell("A1", "Generated in Memory")
      |> Sheet.set_cell("A2", "No file system access")

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, {filename, binary}} = Elixlsx.write_to_memory(workbook, @memory_test_filename)
      assert filename == to_charlist(@memory_test_filename)
      assert is_binary(binary)
      assert byte_size(binary) > 1000  # Should be reasonable size
    end

    test "large dataset performance" do
      # Test with larger dataset to ensure performance is acceptable
      rows = Enum.map(1..1000, fn i ->
        ["Item #{i}", i * 1.5, i * 2.0, i * 3.5]
      end)
      
      sheet = %Sheet{
        name: "Large Dataset",
        rows: [["Item", "Value1", "Value2", "Value3"] | rows]
      }

      workbook = %Workbook{sheets: [sheet]}
      
      # Measure time to ensure reasonable performance
      {time_microseconds, result} = :timer.tc(fn ->
        Elixlsx.write_to(workbook, @test_file_path)
      end)
      
      assert {:ok, _} = result
      assert File.exists?(@test_file_path)
      
      # Should complete within reasonable time (less than 5 seconds for 1000 rows)
      assert time_microseconds < 5_000_000
      
      # Verify file size is appropriate
      file_info = File.stat!(@test_file_path)
      assert file_info.size > 10_000  # Should be at least 10KB for 1000 rows
    end

    test "custom datetime in workbook" do
      # Test custom datetime functionality
      custom_datetime = "2024-01-15T10:30:00Z"
      
      sheet = Sheet.with_name("Custom Date")
      |> Sheet.set_cell("A1", "Created on custom date")
      
      workbook = %Workbook{
        sheets: [sheet],
        datetime: custom_datetime
      }
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "empty and nil cell handling" do
      # Test handling of empty and nil values
      sheet = Sheet.with_name("Empty Cells")
      |> Sheet.set_cell("A1", "Normal Value")
      |> Sheet.set_cell("A2", "")  # Empty string
      |> Sheet.set_cell("A3", nil)  # Nil value
      |> Sheet.set_cell("A4", :empty)  # Explicit empty
      |> Sheet.set_cell("A5", 0)  # Zero value
      |> Sheet.set_cell("A6", false)  # Boolean false
      |> Sheet.set_cell("A7", true)  # Boolean true

      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end
  end

  describe "Excel file validation" do
    test "generated file has correct XLSX structure" do
      # Create a simple workbook
      sheet = Sheet.with_name("Structure Test")
      |> Sheet.set_cell("A1", "Test")
      
      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      
      # Verify it's a valid ZIP file (XLSX is a ZIP archive)
      assert {:ok, file_list} = :zip.list_dir(to_charlist(@test_file_path))
      
      # Check for essential XLSX files
      file_names = Enum.map(file_list, fn 
        {:zip_file, name, _info, _comment, _offset, _comp_size} -> to_string(name)
        {name, _info} -> to_string(name)
      end)
      
      assert "[Content_Types].xml" in file_names
      assert "_rels/.rels" in file_names
      assert "xl/workbook.xml" in file_names
      assert "xl/worksheets/sheet1.xml" in file_names
      assert "xl/sharedStrings.xml" in file_names
      assert "xl/styles.xml" in file_names
    end

    test "file can be opened and read back" do
      # Create a workbook with known data
      test_data = [
        ["Name", "Age", "City"],
        ["Alice", 30, "New York"],
        ["Bob", 25, "Los Angeles"],
        ["Charlie", 35, "Chicago"]
      ]
      
      sheet = %Sheet{
        name: "Test Data",
        rows: test_data
      }
      
      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      
      # Verify the file can be opened as a ZIP and contains expected content
      assert {:ok, files} = :zip.extract(to_charlist(@test_file_path), [:memory])
      
      # Find and verify the worksheet content
      worksheet_content = Enum.find_value(files, fn
        {~c"xl/worksheets/sheet1.xml", content} -> content
        _ -> nil
      end)
      
      assert worksheet_content != nil
      
      # Convert to string and check for our test data
      xml_content = to_string(worksheet_content)
      # The data is stored as references to shared strings, so we need to check the structure
      # Just verify that we have the expected number of rows and cells
      assert String.contains?(xml_content, "<row r=\"1\"")
      assert String.contains?(xml_content, "<row r=\"2\"")
      assert String.contains?(xml_content, "<row r=\"3\"")
      assert String.contains?(xml_content, "<row r=\"4\"")
    end
  end

  describe "Error handling and edge cases" do
    test "handles invalid sheet names gracefully" do
      # Test various invalid sheet name scenarios
      invalid_names = ["", "This is a very long sheet name that exceeds the limit", "Invalid:Chars", "Bad\\Name", "Bad/Name", "Bad?Name", "Bad*Name", "Bad[Name", "Bad]Name"]
      
      Enum.each(invalid_names, fn name ->
        sheet = Sheet.with_name(name)
        workbook = %Workbook{sheets: [sheet]}
        
        assert_raise ArgumentError, fn ->
          Elixlsx.write_to(workbook, @test_file_path)
        end
      end)
    end

    test "handles large cell values" do
      # Test with very long strings
      long_string = String.duplicate("A", 10000)
      
      sheet = Sheet.with_name("Large Values")
      |> Sheet.set_cell("A1", long_string)
      |> Sheet.set_cell("A2", 999999999999999)  # Large number
      
      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
    end

    test "handles many sheets" do
      # Test with multiple sheets (Excel limit is around 255)
      sheets = Enum.map(1..50, fn i ->
        Sheet.with_name("Sheet #{i}")
        |> Sheet.set_cell("A1", "Sheet #{i} Content")
      end)
      
      workbook = %Workbook{sheets: sheets}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      assert File.exists?(@test_file_path)
      
      # Verify file size is reasonable for 50 sheets
      file_info = File.stat!(@test_file_path)
      assert file_info.size > 5000
    end
  end

  describe "Compatibility with Excel applications" do
    test "file format is compatible with modern Excel specifications" do
      # Create a comprehensive workbook that tests various Excel features
      sheet = Sheet.with_name("Compatibility Test")
      |> Sheet.set_cell("A1", "Text Value")
      |> Sheet.set_cell("B1", 42)
      |> Sheet.set_cell("C1", 3.14159)
      |> Sheet.set_cell("D1", true)
      |> Sheet.set_cell("E1", {{2024, 1, 15}, {10, 30, 0}}, datetime: true)
      |> Sheet.set_cell("F1", {:formula, "B1*C1", value: 131.95})
      |> Sheet.set_cell("A2", "Formatted", bold: true, italic: true, color: "#FF0000")
      |> Sheet.set_cell("B2", "Background", bg_color: "#FFFF00")
      |> Sheet.set_cell("C2", "Border", border: [bottom: [style: :thin, color: "#000000"]])
      
      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      
      # Verify the file structure matches Excel expectations
      assert {:ok, file_list} = :zip.list_dir(to_charlist(@test_file_path))
      file_names = Enum.map(file_list, fn 
        {:zip_file, name, _info, _comment, _offset, _comp_size} -> to_string(name)
        {name, _info} -> to_string(name)
      end)
      
      # Check for all required Excel files
      required_files = [
        "[Content_Types].xml",
        "_rels/.rels", 
        "xl/_rels/workbook.xml.rels",
        "xl/workbook.xml",
        "xl/worksheets/sheet1.xml",
        "xl/sharedStrings.xml",
        "xl/styles.xml"
      ]
      
      Enum.each(required_files, fn file ->
        assert file in file_names, "Missing required Excel file: #{file}"
      end)
    end

    test "generated XML is well-formed" do
      # Test that generated XML files are well-formed
      sheet = Sheet.with_name("XML Test")
      |> Sheet.set_cell("A1", "XML & Special <Characters>")
      |> Sheet.set_cell("A2", "Quotes: \"Hello\" & 'World'")
      
      workbook = %Workbook{sheets: [sheet]}
      
      assert {:ok, _} = Elixlsx.write_to(workbook, @test_file_path)
      
      # Extract and validate key XML files
      assert {:ok, files} = :zip.extract(to_charlist(@test_file_path), [:memory])
      
      xml_files = [
        ~c"xl/workbook.xml",
        ~c"xl/worksheets/sheet1.xml", 
        ~c"xl/sharedStrings.xml",
        ~c"xl/styles.xml"
      ]
      
      Enum.each(xml_files, fn xml_file ->
        xml_content = Enum.find_value(files, fn
          {^xml_file, content} -> content
          _ -> nil
        end)
        
        assert xml_content != nil, "Missing XML file: #{xml_file}"
        
        # Verify XML is parseable (basic well-formedness check)
        # Convert to charlist for xmerl_scan
        xml_charlist = if is_binary(xml_content), do: String.to_charlist(xml_content), else: xml_content
        {parsed_xml, _} = :xmerl_scan.string(xml_charlist)
        assert parsed_xml != nil
      end)
    end
  end
end