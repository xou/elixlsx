defmodule Elixlsx.Style.Fill do
  @moduledoc ~S"""
  Fill styling properties.

  Supported formatting properties are:

  - bg_color: (Hex-)String
  """
  alias __MODULE__
  defstruct fg_color: nil

  @type t :: %Fill{
    fg_color: String.t
  }

  @doc ~S"""
  Create a Fill object from a property list.
  """
  def from_props props do
    ft = %Fill{fg_color: props[:bg_color]}

    if ft == %Fill{}, do: nil, else: ft
  end

  defp to_rgb_color(color) do
    # parses a color property and regurns a ARGB code (FFRRGGBB)
    # In the future, this would be the place to support color names such as "red", etc.
    # Also, XLSX has "indexed" colors, see
    # https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.indexedcolors%28v=office.14%29.aspx
    # for a list.
    case String.match?(color, ~r/#[0-9a-fA-F]{6}/) do
      true ->
        "FF" <> (
          color |>
          String.slice(1..-1) |> # remove leading character
          String.capitalize)
      false ->
        raise %ArgumentError{message: "Color values must be in format #rrggbb (hex values), is " <> (inspect color)}
    end
  end

  @spec get_stylexml_entry(Elixlsx.Style.Fill.t) :: String.t
  @doc ~S"""
  Create a <fill /> entry from a Fill struct.
  """
  def get_stylexml_entry(fill) do
    fg_color = if fill.fg_color do "<patternFill patternType=\"solid\"><fgColor rgb=\"#{to_rgb_color(fill.fg_color)}\" /></patternFill>" else "" end

    "<fill>#{fg_color}</fill>"
  end
end