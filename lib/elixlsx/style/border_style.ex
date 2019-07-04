defmodule Elixlsx.Style.BorderStyle do
  @moduledoc ~S"""
  Border styling properties
  """
  alias __MODULE__
  alias Elixlsx.Style.Border

  defstruct left: nil,
            right: nil,
            top: nil,
            bottom: nil,
            diagonal: nil,
            diagonal_up: false,
            diagonal_down: false

  @type t :: %BorderStyle{
          left: Border.t(),
          right: Border.t(),
          top: Border.t(),
          bottom: Border.t(),
          diagonal: Border.t(),
          diagonal_up: boolean,
          diagonal_down: boolean
        }

  def from_props(props) do
    left = Border.from_props(props[:left], :left)
    right = Border.from_props(props[:right], :right)
    top = Border.from_props(props[:top], :top)
    bottom = Border.from_props(props[:bottom], :bottom)
    diagonal = Border.from_props(props[:diagonal], :diagonal)

    %BorderStyle{
      left: left,
      right: right,
      top: top,
      bottom: bottom,
      diagonal: diagonal,
      diagonal_up: !!props[:diagonal_up],
      diagonal_down: !!props[:diagonal_down]
    }
  end

  @spec get_border_style_entry(BorderStyle.t()) :: String.t()
  @doc ~S"""
   Generate xml entry for border group

  ## Examples

    iex> Elixlsx.Style.BorderStyle.get_border_style_entry Elixlsx.Style.BorderStyle.from_props top: [style: :dash_dot, color: "#eeccaa"]
    "<border diagonalUp=\"false\" diagonalDown=\"false\">\n  <left></left><right></right><top style=\"dashDot\"><color rgb=\"FFEECCAA\" /></top><bottom></bottom><diagonal></diagonal>\n</border>\n"

  """
  def get_border_style_entry(border) do
    diagonal_up = if border.diagonal_up, do: "true", else: "false"
    diagonal_down = if border.diagonal_down, do: "true", else: "false"
    left = Border.get_border_entry(border.left)
    right = Border.get_border_entry(border.right)
    top = Border.get_border_entry(border.top)
    bottom = Border.get_border_entry(border.bottom)

    # ignore diagonal border, if none of flags set
    diagonal =
      if border.diagonal_up or border.diagonal_down do
        Border.get_border_entry(border.diagonal)
      else
        "<diagonal></diagonal>"
      end

    """
    <border diagonalUp="#{diagonal_up}" diagonalDown="#{diagonal_down}">
      #{left}#{right}#{top}#{bottom}#{diagonal}
    </border>
    """
  end
end
