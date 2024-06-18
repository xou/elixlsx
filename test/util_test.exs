defmodule UtilTest do
  use ExUnit.Case

  alias Elixlsx.Image
  alias Elixlsx.Util

  test "width_to_px" do
    assert Util.width_to_px(1, %Image{}) == 12
    assert Util.width_to_px(1, %Image{char: 10}) == 15
  end
end
