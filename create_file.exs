#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

data = %Sheet{name: 'First', rows: [[1,2,3],[4,5,6]]}
Elixlsx.write_to(data, "empty.xlsx")
