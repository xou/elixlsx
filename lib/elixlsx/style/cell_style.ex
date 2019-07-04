defmodule Elixlsx.Style.CellStyle do
  alias __MODULE__
  alias Elixlsx.Style.NumFmt
  alias Elixlsx.Style.Font
  alias Elixlsx.Style.Fill
  alias Elixlsx.Style.BorderStyle

  defstruct font: nil, fill: nil, numfmt: nil, border: nil

  @type t :: %CellStyle{
          font: Font.t(),
          fill: Fill.t(),
          numfmt: NumFmt.t(),
          border: BorderStyle.t()
        }

  def from_props(props) do
    font = Font.from_props(props)
    fill = Fill.from_props(props)
    numfmt = NumFmt.from_props(props)
    border = BorderStyle.from_props(props[:border])

    %CellStyle{font: font, fill: fill, numfmt: numfmt, border: border}
  end

  def is_date?(cellstyle) do
    cond do
      is_nil(cellstyle) -> false
      is_nil(cellstyle.numfmt) -> false
      true -> NumFmt.is_date?(cellstyle.numfmt)
    end
  end
end
