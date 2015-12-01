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
  strike: false, size: nil

  @type t :: %Font{
    bold: boolean,
    italic: boolean,
    underline: boolean,
    strike: boolean,
    size: pos_integer
  }

  @doc ~S"""
  Create a Font object from a property list.
  """
  def from_props props do
    ft = %Font{bold: !!props[:bold],
               italic: !!props[:italic],
               underline: !!props[:underline],
               strike: !!props[:strike],
               size: props[:size]
              }

    if ft == %Font{}, do: nil, else: ft
  end


  @spec get_stylexml_entry(Elixlsx.Style.Font.t) :: String.t
  @doc ~S"""
  Create a <font /> entry from a Font struct.
  """
  def get_stylexml_entry(font) do
    bold = if font.bold do "<b val=\"1\"/>" else "" end
    italic = if font.italic do "<i val=\"1\"/>" else "" end
    # TODO underline doesn't work.
    underline = if font.underline do "<u val=\"1\"/>" else "" end
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

    "<font>#{bold}#{italic}#{underline}#{strike}#{size}</font>"
  end
end
