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
workbook = %Workbook{sheets: [sheet, sheet2]}
Elixlsx.write_to(workbook, "empty.xlsx")
