#!/bin/bash

test_dir=`dirname $0`
pushd $test_dir/../
mix
elixir -pa _build/dev/lib/elixlsx/ebin example.exs
libreoffice --convert-to empty.pdf empty.xlsx
rc=$?
popd

exit $rc

