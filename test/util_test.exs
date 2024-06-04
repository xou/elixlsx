use PropCheck

defmodule UtilTest do
  use ExUnit.Case, async: false

  alias Elixlsx.Util

  property "Util.encode_col reverses decode_col", [:verbose] do
    forall x <- non_neg_integer() do
      assert Util.decode_col(Util.encode_col(x)) == x
    end
  end
end
