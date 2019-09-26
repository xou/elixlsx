defmodule ExCheck.UtilTest do
  use ExUnit.Case, async: false
  use ExCheck

  alias Elixlsx.Util

  property :enc_dec do
    for_all x in such_that(x in int() when x >= 0) do
      implies x >= 0 do
        Util.decode_col(Util.encode_col(x)) == x
      end
    end
  end
end
