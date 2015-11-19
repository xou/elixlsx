#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

sheet = %Sheet{name: 'First', rows: [[1,2,3],[4,5,6]]}
workbook = %Workbook{sheets: [sheet]}
Elixlsx.write_to(workbook, "empty.xlsx")
