defmodule Elixlsx.Compiler.FillDB do
  alias __MODULE__
  alias Elixlsx.Style.Fill
  alias Elixlsx.Compiler.DBUtil

  defstruct fills: %{}, element_count: 0

  @type t :: %FillDB{
          fills: %{Fill.t() => pos_integer},
          element_count: non_neg_integer
        }

  @spec register_fill(FillDB.t(), Fill.t()) :: FillDB.t()
  def register_fill(filldb, fill) do
    case Map.fetch(filldb.fills, fill) do
      :error ->
        %FillDB{
          fills: Map.put(filldb.fills, fill, filldb.element_count + 2),
          element_count: filldb.element_count + 1
        }

      {:ok, _} ->
        filldb
    end
  end

  def get_id(filldb, fill) do
    case Map.fetch(filldb.fills, fill) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for FillDB.get_id: " <> inspect(fill)}

      {:ok, id} ->
        id
    end
  end

  def id_sorted_fills(filldb), do: DBUtil.id_sorted_values(filldb.fills)
end
