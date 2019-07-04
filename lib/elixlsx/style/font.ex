defmodule Elixlsx.Style.Font do
  @moduledoc ~S"""
  Font styling properties.

  Supported formatting properties are:

  - bold: boolean
  - italic: boolean
  - underline: boolean
  - strike: boolean
  - size: pos_integer
  - color: (Hex-)String
  - wrap_text: boolean
  - align_horizontal: atom (:left, :right, :center, :justify, :general, :fill)
  - align_vertical: atom (:top, :bottom, :center)
  - font: String
  """
  import Elixlsx.Color, only: [to_rgb_color: 1]
  alias __MODULE__

  defstruct bold: false,
            italic: false,
            underline: false,
            strike: false,
            size: nil,
            font: nil,
            color: nil,
            wrap_text: false,
            align_horizontal: nil,
            align_vertical: nil

  @type t :: %Font{
          bold: boolean,
          italic: boolean,
          underline: boolean,
          strike: boolean,
          size: pos_integer,
          color: String.t(),
          wrap_text: boolean,
          align_horizontal: atom,
          align_vertical: atom,
          font: String.t()
        }

  @doc ~S"""
  Create a Font object from a property list.
  """
  def from_props(props) do
    ft = %Font{
      bold: !!props[:bold],
      italic: !!props[:italic],
      underline: !!props[:underline],
      strike: !!props[:strike],
      size: props[:size],
      color: props[:color],
      wrap_text: !!props[:wrap_text],
      align_horizontal: props[:align_horizontal],
      align_vertical: props[:align_vertical],
      font: props[:font]
    }

    if ft == %Font{}, do: nil, else: ft
  end

  @spec get_stylexml_entry(Elixlsx.Style.Font.t()) :: String.t()
  @doc ~S"""
  Create a <font /> entry from a Font struct.
  """
  def get_stylexml_entry(font) do
    bold =
      if font.bold do
        "<b val=\"1\"/>"
      else
        ""
      end

    italic =
      if font.italic do
        "<i val=\"1\"/>"
      else
        ""
      end

    # TODO: Add more underline properties, see here:
    # http://webapp.docx4java.org/OnlineDemo/ecma376/SpreadsheetML/ST_UnderlineValues.html
    underline =
      if font.underline do
        "<u val=\"single\"/>"
      else
        ""
      end

    strike =
      if font.strike do
        "<strike val=\"1\"/>"
      else
        ""
      end

    size =
      if font.size do
        case is_number(font.size) do
          true ->
            "<sz val=\"#{font.size}\"/>"

          false ->
            raise %ArgumentError{message: "Invalid font size: " <> inspect(font.size)}
        end
      else
        ""
      end

    color =
      if font.color do
        "<color rgb=\"#{to_rgb_color(font.color)}\" />"
      else
        ""
      end

    font_name =
      if font.font do
        "<name val=\"#{font.font}\" />"
      else
        ""
      end

    "<font>#{bold}#{italic}#{underline}#{strike}#{size}#{font_name}#{color}</font>"
  end
end
