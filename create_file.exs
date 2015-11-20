#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

sheet = %Sheet{name: 'First', rows: [[1,2,3],[4,5,6]]}
sheet2 = %Sheet{name: 'Second', rows: [[1,2,3,4,5],[1,2]]}
workbook = %Workbook{sheets: [sheet, sheet2]}
Elixlsx.write_to(workbook, "empty.xlsx")
