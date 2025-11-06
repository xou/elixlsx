defmodule ElixlsxPerformanceTest do
  use ExUnit.Case

  alias Elixlsx.{Sheet, Workbook}

  @moduletag :performance

  describe "Performance validation" do
    test "small workbook generation performance" do
      # Create a small workbook with basic data
      sheet = Sheet.with_name("Performance Test")
        |> Sheet.set_cell("A1", "Name")
        |> Sheet.set_cell("B1", "Value")
        |> Sheet.set_cell("A2", "Test")
        |> Sheet.set_cell("B2", 123.45)

      workbook = %Workbook{sheets: [sheet]}

      # Measure time for compilation and writing
      {time_microseconds, _result} = :timer.tc(fn ->
        Elixlsx.write_to(workbook, "test_small_performance.xlsx")
      end)

      # Clean up
      File.rm("test_small_performance.xlsx")

      # Convert to milliseconds for easier reading
      time_ms = time_microseconds / 1000

      # Assert reasonable performance (should complete in under 100ms for small file)
      assert time_ms < 100, "Small workbook generation took #{time_ms}ms, expected < 100ms"

      IO.puts("Small workbook generation: #{Float.round(time_ms, 2)}ms")
    end

    test "medium workbook generation performance" do
      # Create a medium-sized workbook with multiple sheets and formatting
      sheet1 = Sheet.with_name("Data Sheet")
        |> add_test_data(100)  # 100 rows of data
        |> Sheet.set_col_width("A", 15.0)
        |> Sheet.set_col_width("B", 12.0)

      sheet2 = Sheet.with_name("Formatted Sheet")
        |> add_formatted_data(50)  # 50 rows with formatting

      workbook = %Workbook{sheets: [sheet1, sheet2]}

      # Measure time for compilation and writing
      {time_microseconds, _result} = :timer.tc(fn ->
        Elixlsx.write_to(workbook, "test_medium_performance.xlsx")
      end)

      # Clean up
      File.rm("test_medium_performance.xlsx")

      # Convert to milliseconds
      time_ms = time_microseconds / 1000

      # Assert reasonable performance (should complete in under 1000ms for medium file)
      assert time_ms < 1000, "Medium workbook generation took #{time_ms}ms, expected < 1000ms"

      IO.puts("Medium workbook generation: #{Float.round(time_ms, 2)}ms")
    end

    test "large workbook generation performance" do
      # Create a larger workbook to test scalability
      sheet1 = Sheet.with_name("Large Data")
        |> add_test_data(1000)  # 1000 rows of data

      sheet2 = Sheet.with_name("More Data")
        |> add_test_data(500)

      workbook = %Workbook{sheets: [sheet1, sheet2]}

      # Measure time for compilation and writing
      {time_microseconds, _result} = :timer.tc(fn ->
        Elixlsx.write_to(workbook, "test_large_performance.xlsx")
      end)

      # Clean up
      File.rm("test_large_performance.xlsx")

      # Convert to milliseconds
      time_ms = time_microseconds / 1000

      # Assert reasonable performance (should complete in under 5000ms for large file)
      assert time_ms < 5000, "Large workbook generation took #{time_ms}ms, expected < 5000ms"

      IO.puts("Large workbook generation: #{Float.round(time_ms, 2)}ms")
    end

    test "memory usage validation" do
      # Test that memory usage doesn't grow excessively
      initial_memory = :erlang.memory(:total)

      # Create and generate multiple workbooks
      Enum.each(1..10, fn i ->
        sheet = Sheet.with_name("Sheet #{i}")
          |> add_test_data(100)

        workbook = %Workbook{sheets: [sheet]}
        filename = "test_memory_#{i}.xlsx"
        
        Elixlsx.write_to(workbook, filename)
        File.rm(filename)
      end)

      # Force garbage collection
      :erlang.garbage_collect()
      
      final_memory = :erlang.memory(:total)
      memory_growth = final_memory - initial_memory

      # Memory growth should be reasonable (less than 10MB for this test)
      max_growth = 10 * 1024 * 1024  # 10MB in bytes
      assert memory_growth < max_growth, 
        "Memory grew by #{memory_growth} bytes, expected < #{max_growth} bytes"

      IO.puts("Memory growth: #{Float.round(memory_growth / 1024 / 1024, 2)}MB")
    end
  end

  # Helper function to add test data to a sheet
  defp add_test_data(sheet, row_count) do
    Enum.reduce(1..row_count, sheet, fn i, acc_sheet ->
      acc_sheet
      |> Sheet.set_cell("A#{i}", "Item #{i}")
      |> Sheet.set_cell("B#{i}", i * 1.5)
      |> Sheet.set_cell("C#{i}", "Description for item #{i}")
    end)
  end

  # Helper function to add formatted test data
  defp add_formatted_data(sheet, row_count) do
    Enum.reduce(1..row_count, sheet, fn i, acc_sheet ->
      acc_sheet
      |> Sheet.set_cell("A#{i}", "Formatted #{i}", bold: true, color: "#0066cc")
      |> Sheet.set_cell("B#{i}", i * 2.5, num_format: "0.00", bg_color: "#f0f0f0")
      |> Sheet.set_cell("C#{i}", "Status #{i}", italic: true)
    end)
  end
end