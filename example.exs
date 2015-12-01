#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

alias Elixlsx.Sheet
alias Elixlsx.Workbook


sheet1 = Sheet.with_name("First")
# Set cell B2 to the string "Hi". :)
         |> Sheet.set_cell("B2", "Hi")
# Optionally, set font properties:
         |> Sheet.set_cell("B3", "Hello World", bold: true, underline: true)
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

workbook = %Workbook{sheets: [sheet1]}

# it is also possible to add a custom "created" date to workbook, otherwise,
# the current date is used.

workbook = %Workbook{workbook | datetime: "2015-12-01T13:40:59Z"}

# It is also possible to create a sheet as a list of rows:
sheet2 = %Sheet{name: 'Third', rows: [[1,2,3,4,5],
                                       [1,2],
                                       ["hello", "world"]]}
workbook = Workbook.append_sheet(workbook, sheet2)

# For the list of rows approach, cells with properties can be encoded by using a
# list with the value at the head and the properties in the tail:
sheet3 = %Sheet{name: "Second", rows:
  [[1,2,3],
   [4,5,6, ["goat", bold: true]],
   [["Bold", bold: true], ["Italic", italic: true], ["Underline", underline: true], ["Strike!", strike: true],
    ["Large", size: 22]],
# Unicode should work as well:
   [["Müłti", bold: true, italic: true, underline: true, strike: true]]
  ]}

Workbook.insert_sheet(workbook, sheet3, 1)
|> Elixlsx.write_to("empty.xlsx")
