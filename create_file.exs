#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

sheet = %Sheet{name: 'First', rows: 
  [[1,2,3],
   [4,5,6, ["goat", bold: true]],
   [["Bold", bold: true], ["Italic", italic: true], ["Underline", underline: true], ["Strike!", strike: true],
    ["Large", size: 22]],
   [["Müłti", bold: true, italic: true, underline: true, strike: true]]
  ]}
sheet2 = %Sheet{name: 'Second', rows: [[1,2,3,4,5],[1,2], ["hello", "goat", "world"]]}
sheet3 = Sheet.with_name("Third")
         |> Sheet.set_cell("B3", "Hello World", bold: true, underline: true)
         |> Sheet.set_cell("A1", Elixlsx.Util.to_excel_datetime(1448882362), yyyymmdd: true)
         |> Sheet.set_cell("A2", Elixlsx.Util.to_excel_datetime({{2015, 11, 30}, {21, 20, 38}}), datetime: true)
         |> Sheet.set_cell("A3", 123.4, num_format: "0.00")
         |> Sheet.set_cell("A4", {{2015, 11, 30}, {21, 20, 38}}, datetime: true)
         |> Sheet.set_cell("A5", 1448882362, yyyymmdd: true)

workbook = %Workbook{sheets: [sheet, sheet2, sheet3]}
Elixlsx.write_to(workbook, "empty.xlsx")
