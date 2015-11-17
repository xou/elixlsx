#!/usr/bin/elixir -pa _build/dev/lib/elixlsx/ebin/

require Elixlsx

Elixlsx.write_to([], "empty.xlsx")
