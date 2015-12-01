defmodule Elixlsx.Style.CellStyle do
  alias __MODULE__
  alias Elixlsx.Style.NumFmt
  alias Elixlsx.Style.Font

  defstruct font: nil, numfmt: nil

  @type t :: %CellStyle{
    font: Font.t
  }


  def from_props props do
    font = Font.from_props props
    numfmt = NumFmt.from_props props

    %CellStyle{font: font,
               numfmt: numfmt}
  end

  def is_date?(cellstyle) do
    cond do
      is_nil(cellstyle) -> false
      is_nil(cellstyle.numfmt) -> false
      true -> NumFmt.is_date? cellstyle.numfmt
    end
  end
end
