defmodule Elixlsx.Style.Border do
  @moduledoc ~S"""
  Border styling properties
  - type: atom (:left, :right, :top, :bottom, :diagonal)
  - style: atom (:solid, :dotted, :dashed, :double,
              :dash_dot, :dash_dot_dot, :thin, :medium, :thick)
  - color: (Hex-)String
  """
  alias __MODULE__

  defstruct type: nil,
            style: nil,
            color: nil

  @type t :: %Border{
          type: atom,
          style: atom,
          color: String.t()
        }

  def from_props(props, type) do
    %Border{
      type: type,
      style: props[:style],
      color: props[:color]
    }
  end

  @spec border_name(atom) :: String.t()
  defp border_name(type) do
    if type in [:left, :right, :top, :bottom, :diagonal] do
      Atom.to_string(type)
    else
      raise %ArgumentError{
        message: "Invalid border type: " <> inspect(type)
      }
    end
  end

  @spec style_name(atom) :: String.t()
  defp style_name(style) do
    case style do
      :thin ->
        "thin"

      :medium ->
        "medium"

      :thick ->
        "thick"

      :dotted ->
        "dotted"

      :dashed ->
        "dashed"

      :double ->
        "double"

      :dash_dot ->
        "dashDot"

      :dash_dot_dot ->
        "dashDotDot"

      _ ->
        raise %ArgumentError{
          message: "Invalid border style: " <> inspect(style)
        }
    end
  end

  @spec border_style(atom) :: String.t()
  defp border_style(nil), do: ""
  defp border_style(style), do: " style=\"#{style_name(style)}\""

  @spec border_color(String.t()) :: String.t()
  defp border_color(nil), do: ""

  defp border_color(color) do
    value = Elixlsx.Color.to_rgb_color(color)
    "<color rgb=\"#{value}\" />"
  end

  @spec get_border_entry(Border.t()) :: String.t()
  @doc ~S"""
   Create an entity for border

  ## Examples

    iex> Elixlsx.Style.Border.get_border_entry %Elixlsx.Style.Border{type: :top, style: :dotted}
    "<top style=\"dotted\"></top>"

  """
  def get_border_entry(border) do
    name = border_name(border.type)
    style = border_style(border.style)
    color = border_color(border.color)

    "<#{name}#{style}>#{color}</#{name}>"
  end
end
