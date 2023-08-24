defmodule ExCheck.UtilTest do
  use ExUnit.Case
  use ExCheck

  alias Elixlsx.Util
  alias Elixlsx.Sheet

  doctest Util, import: true

  property :enc_dec do
    for_all x in such_that(x in int() when x >= 0) do
      implies x >= 0 do
        Util.decode_col(Util.encode_col(x)) == x
      end
    end
  end

  test "width_from_col_range" do
    sheet = %Sheet{col_widths: %{1 => 1, 2 => 2}}
    assert Util.width_from_col_range(sheet, 0, 2) == 11.43
  end

  test "height_from_row_range" do
    sheet = %Sheet{row_heights: %{1 => 1, 2 => 2}}
    assert Util.height_from_row_range(sheet, 0, 2) == 18
  end

  test "width_to_px" do
    sheet = %Sheet{}
    assert Util.width_to_px(sheet, 1) == 12

    sheet = %Sheet{max_char_width: 10}
    assert Util.width_to_px(sheet, 1) == 15
  end

  test "width_to_emu" do
    sheet = %Sheet{emu: 10}
    assert Util.width_to_emu(sheet, 1) == 120

    sheet = %Sheet{emu: 10, max_char_width: 10}
    assert Util.width_to_emu(sheet, 1) == 150
  end

  test "height_to_px" do
    assert Util.height_to_px(1) == 1.3333333333333333
  end

  test "height_to_emu" do
    sheet = %Sheet{emu: 10}
    assert Util.height_to_emu(sheet, 1) == 13
  end

  test "px_to_width" do
    sheet = %Sheet{}
    assert Util.px_to_width(sheet, 1) == 0.08333333333333333

    sheet = %Sheet{max_char_width: 10}
    assert Util.px_to_width(sheet, 1) == 0.06666666666666667
  end

  test "px_to_height" do
    assert Util.px_to_height(1) == 0.75
  end

  test "px_to_col_span :left" do
    sheet =
      %Sheet{max_char_width: 10}
      |> Sheet.set_col_width("A", "100px")
      |> Sheet.set_col_width("B", "100px")
      |> Sheet.set_col_width("C", "100px")

    assert Util.px_to_col_span_from_left(sheet, 0, 25) == {0, 0, {0, 25}}
    assert Util.px_to_col_span_from_left(sheet, 0, 100) == {0, 0, {0, 100}}
    assert Util.px_to_col_span_from_left(sheet, 0, 200) == {0, 1, {0, 100}}
    assert Util.px_to_col_span_from_left(sheet, 0, 201) == {0, 2, {1, 100}}
    assert Util.px_to_col_span_from_left(sheet, 0, 300) == {0, 2, {0, 100}}
  end

  test "px_to_col_span :right" do
    sheet =
      %Sheet{max_char_width: 10}
      |> Sheet.set_col_width("A", "100px")
      |> Sheet.set_col_width("B", "100px")
      |> Sheet.set_col_width("C", "100px")

    assert Util.px_to_col_span_from_right(sheet, 2, 25) == {2, 2, {75, 100}}
    assert Util.px_to_col_span_from_right(sheet, 2, 100) == {2, 2, {0, 100}}
    assert Util.px_to_col_span_from_right(sheet, 2, 200) == {1, 2, {0, 100}}
    assert Util.px_to_col_span_from_right(sheet, 2, 201) == {0, 2, {99, 100}}
    assert Util.px_to_col_span_from_right(sheet, 2, 300) == {0, 2, {0, 100}}
  end

  test "px_to_row_span" do
    sheet = %Sheet{max_char_width: 8}
    assert Util.px_to_row_span(sheet, 0, 20) == {0, 0, 20}
    assert Util.px_to_row_span(sheet, 0, 21) == {0, 1, 1}
    assert Util.px_to_row_span(sheet, 0, 80) == {0, 3, 20}
  end
end
