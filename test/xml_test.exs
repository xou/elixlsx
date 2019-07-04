defmodule XMLTest do
  use ExUnit.Case

  test "valid? allows normal string" do
    assert XML.valid?("Hello World & Goodbye Cruel World")
  end

  test "valid? rejects invalid xml characters" do
    refute XML.valid?(<<31>>)
  end
end
