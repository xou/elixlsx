defmodule ExCheck.UtilTest do
  use ExUnit.Case
  use ExCheck

  alias Elixlsx.Image
  alias Elixlsx.Util

  doctest Util, import: true

  property :enc_dec do
    for_all x in such_that(x in int() when x >= 0) do
      implies x >= 0 do
        Util.decode_col(Util.encode_col(x)) == x
      end
    end
  end

  test "width_to_px" do
    assert Util.width_to_px(1, %Image{}) == 12
    assert Util.width_to_px(1, %Image{char: 10}) == 15
  end
end
