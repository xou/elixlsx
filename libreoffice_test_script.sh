#!/bin/bash

elixir -pa _build/dev/lib/elixlsx/ebin create_file.exs
libreoffice --convert-to empty.pdf empty.xlsx
exit $?

