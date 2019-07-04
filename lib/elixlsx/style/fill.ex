defmodule Elixlsx.Style.Fill do
  @moduledoc ~S"""
  Fill styling properties.

  Supported formatting properties are:

  - bg_color: (Hex-)String
  """
  alias __MODULE__
  import Elixlsx.Color, only: [to_rgb_color: 1]
  defstruct fg_color: nil

  @type t :: %Fill{
          fg_color: String.t()
        }

  @doc ~S"""
  Create a Fill object from a property list.
  """
  def from_props(props) do
    ft = %Fill{fg_color: props[:bg_color]}

    if ft == %Fill{}, do: nil, else: ft
  end

  @spec get_stylexml_entry(Elixlsx.Style.Fill.t()) :: String.t()
  @doc ~S"""
  Create a <fill /> entry from a Fill struct.
  """
  def get_stylexml_entry(fill) do
    fg_color =
      if fill.fg_color do
        "<patternFill patternType=\"solid\"><fgColor rgb=\"#{to_rgb_color(fill.fg_color)}\" /></patternFill>"
      else
        ""
      end

    "<fill>#{fg_color}</fill>"
  end
end
