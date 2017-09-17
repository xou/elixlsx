#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

alias Elixlsx.Sheet
alias Elixlsx.Workbook


sheet1 = Sheet.with_name("First")
# Set cell B2 to the string "Hi". :)
         |> Sheet.set_cell("B2", "Hi")
# Optionally, set font properties:
         |> Sheet.set_cell("B3", "Hello World", bold: true, underline: true, color: "#ffaa00")
# Set background color
         |> Sheet.set_cell("B4", "Background color \\o/", bg_color: "#ffff00")
# Number formatting can be applied like this:
         |> Sheet.set_cell("A1", 123.4, num_format: "0.00")
# Two date formats are accepted, erlang's :calendar format and UNIX timestamps.
# the datetime: true parameter automatically applies conversion to Excels internal format.
         |> Sheet.set_cell("A2", {{2015, 11, 30}, {21, 20, 38}}, datetime: true)
         |> Sheet.set_cell("A3", 1448882362, datetime: true)
# datetime: true ouputs date and time, yyyymmdd limits the output to just the date
         |> Sheet.set_cell("A4", 1448882362, yyyymmdd: true)
# make some room in the first column, otherwise the date will only show up as ###
         |> Sheet.set_col_width("A", 18.0)
# Cell borders
         |> Sheet.set_cell("A5", "Double border", border: [bottom: [style: :double, color: "#cc3311"]])
# Formatting with empty content
         |> Sheet.set_cell("A5", :empty, bg_color: "#ffff00", border: [bottom: [style: :double, color: "#cc3311"]])
# Formula
         |> Sheet.set_cell("E1", 1.2, num_format: "0.00")
         |> Sheet.set_cell("E2", 2, num_format: "0.00")
         |> Sheet.set_cell("E3", 3.5, num_format: "0.00")
         |> Sheet.set_cell("E4", 4, num_format: "0.00")
         |> Sheet.set_cell("E5", 5, num_format: "0.00")
 # Show `value` if there is no calculated value (example: Mac OS preview):
         |> Sheet.set_cell("E6", {:formula, "SUM(E1:E5)", value: 15.70}, num_format: "0.00", bold: true)
         |> Sheet.set_cell("F1", {:formula, "NOW()"}, num_format: "yyyy-mm-dd hh:MM:ss")
         |> Sheet.set_col_width("F", 18.0)

workbook = %Workbook{sheets: [sheet1]}

# it is also possible to add a custom "created" date to workbook, otherwise,
# the current date is used.

workbook = %Workbook{workbook | datetime: "2015-12-01T13:40:59Z"}

# It is also possible to create a sheet as a list of rows:
sheet2 = %Sheet{name: "Third", rows: [[1,2,3,4,5],
                                       [1,2],
                                       ["increased row height"],
                                       ["hello", "world"]]}
         |> Sheet.set_row_height(3, 40)

workbook = Workbook.append_sheet(workbook, sheet2)

# For the list of rows approach, cells with properties can be encoded by using a
# list with the value at the head and the properties in the tail:
sheet3 = %Sheet{name: "Second", rows:
  [[1,2,3],
   [4,5,6, ["goat", bold: true]],
   [["Bold", bold: true], ["Italic", italic: true], ["Underline", underline: true], ["Strike!", strike: true],
    ["Large", size: 22]],
   # wrap_text makes text wrap, but it does not increase the row height
   # (see row_heights below).
   [["This is a cell with quite a bit of text.", wrap_text: true],
   # make some vertical alignment
    ["Top", align_vertical: :top],
    # also set font name
    ["Middle", align_vertical: :center, font: "Courier New"],
    ["Bottom", align_vertical: :bottom]],
# Unicode should work as well:
   [["Müłti", bold: true, italic: true, underline: true, strike: true]],
# Change horizontal alignment
   [["left", align_horizontal: :left], ["right", align_horizontal: :right],
    ["center", align_horizontal: :center], ["justify", align_horizontal: :justify],
    ["general", align_horizontal: :general], ["fill", align_horizontal: :fill]]
  ],
  row_heights: %{4 => 60}}

# Insert sheet3 as the second sheet:
workbook = Workbook.insert_sheet(workbook, sheet3, 1)

# It is possible to merge blocks of cells. The top-left cells' value will be retained
# in the block, all others will be dropped.
sheet4 = %Sheet{name: "Merged Cells",
                rows: List.duplicate(["A", "B", "C", "D", "E"], 5),
                merge_cells: [{"A1", "A3"}, {"C1", "E1"}, {"C3", "E5"}]}

# You can set pane freeze.
sheet4
|> Sheet.set_pane_freeze(1, 1) # first row and first column frozen
|> Sheet.set_pane_freeze(2, 1) # first and second row and first column frozen
|> Sheet.remove_pane_freeze # unfreeze pane

sheet5 = %Sheet{name: "No gridlines shown", show_grid_lines: false}
         |> Sheet.set_at(0, 0, "Just this cell")

Workbook.append_sheet(workbook, sheet4)
|> Workbook.append_sheet(sheet5)
|> Elixlsx.write_to("example.xlsx")
