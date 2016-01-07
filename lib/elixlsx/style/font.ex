defmodule Elixlsx.Style.Font do
  @moduledoc ~S"""
  Font styling properties.

  Supported formatting properties are:

  - bold: boolean
  - italic: boolean
  - underline: boolean
  - strike: boolean
  - size: pos_integer

  """
  alias __MODULE__
  defstruct bold: false, italic: false, underline: false,
  strike: false, size: nil, color: nil

  @type t :: %Font{
    bold: boolean,
    italic: boolean,
    underline: boolean,
    strike: boolean,
    size: pos_integer,
    color: String.t
  }


  @doc ~S"""
  Create a Font object from a property list.
  """
  def from_props props do
    ft = %Font{bold: !!props[:bold],
               italic: !!props[:italic],
               underline: !!props[:underline],
               strike: !!props[:strike],
               size: props[:size],
               color: props[:color]
              }

    if ft == %Font{}, do: nil, else: ft
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
        raise %ArgumentError{message: "Font color must be in format #rrggbb (hex values), is " <> (inspect color)}
    end
  end

  @spec get_stylexml_entry(Elixlsx.Style.Font.t) :: String.t
  @doc ~S"""
  Create a <font /> entry from a Font struct.
  """
  def get_stylexml_entry(font) do
    bold = if font.bold do "<b val=\"1\"/>" else "" end
    italic = if font.italic do "<i val=\"1\"/>" else "" end
    # TODO: Add more underline properties, see here:
    # http://webapp.docx4java.org/OnlineDemo/ecma376/SpreadsheetML/ST_UnderlineValues.html
    underline = if font.underline do "<u val=\"single\"/>" else "" end
    strike = if font.strike do "<strike val=\"1\"/>" else "" end
    size = if font.size do
      case is_number(font.size) do
        true ->
          "<sz val=\"#{font.size}\"/>"
        false ->
          raise %ArgumentError{message: "Invalid font size: " <> (inspect font.size)}
      end
    else
      ""
    end

    color = if font.color do "<color rgb=\"#{to_rgb_color(font.color)}\" />" else "" end

    "<font>#{bold}#{italic}#{underline}#{strike}#{size}#{color}</font>"
  end
end
