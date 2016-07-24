defmodule Elixlsx.Compiler.FillDB do
  alias __MODULE__
  alias Elixlsx.Style.Fill

  defstruct fills: %{}, element_count: 0

  @type t :: %FillDB {
    fills: %{Fill.t => pos_integer},
    element_count: non_neg_integer
  }

  @spec register_fill(FillDB.t, Fill.t) :: FillDB.t
  def register_fill(filldb, fill) do
    case Dict.fetch(filldb.fills, fill) do
      :error -> %FillDB{fills: Dict.put(filldb.fills, fill, filldb.element_count + 1),
                       element_count: filldb.element_count + 1}
      {:ok, _} -> filldb
    end
  end

  def get_id(filldb, fill) do
    case Dict.fetch(filldb.fills, fill) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for FillDB.get_id: " <> inspect(fill)}
      {:ok, id} ->
        id
    end
  end

  def id_sorted_fills(filldb) do
    filldb.fills
    |> Enum.map(fn ({k, v}) -> {v, k} end)
    |> Enum.sort
    |> Dict.values
  end
end
