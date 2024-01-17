defmodule Elixlsx.Color do
  @doc ~S"""
  Parses a color property and regurns a ARGB code (FFRRGGBB).

  In the future, this would be the place to support color names such as "red", etc.
  Also, XLSX has "indexed" colors, see
  https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors%28v=office.14%29.aspx
  for a list.

  ## Examples

      iex> Elixlsx.Color.to_rgb_color("#aa5533")
      "FFAA5533"

  """
  def to_rgb_color(color) do
    case String.match?(color, ~r/#[0-9a-fA-F]{6}/) do
      true ->
        "FF" <>
          (color
           # remove leading character
           |> String.slice(1..-1//1)
           |> String.upcase())

      false ->
        raise %ArgumentError{
          message: "Font color must be in format #rrggbb (hex values), is " <> inspect(color)
        }
    end
  end
end
