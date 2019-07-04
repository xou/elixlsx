defmodule Elixlsx.Compiler.BorderStyleDB do
  alias __MODULE__
  alias Elixlsx.Style.BorderStyle
  alias Elixlsx.Compiler.DBUtil

  defstruct borders: %{}, element_count: 0

  @type t :: %BorderStyleDB{
          borders: %{BorderStyle.t() => pos_integer},
          element_count: non_neg_integer
        }

  def register_border(borderstyledb, border) do
    case Map.fetch(borderstyledb.borders, border) do
      :error ->
        %BorderStyleDB{
          borders: Map.put(borderstyledb.borders, border, borderstyledb.element_count + 1),
          element_count: borderstyledb.element_count + 1
        }

      {:ok, _} ->
        borderstyledb
    end
  end

  def get_id(borderstyledb, border) do
    case Map.fetch(borderstyledb.borders, border) do
      :error ->
        raise %ArgumentError{
          message: "Invalid key provided for BorderStyleDB.get_id: " <> inspect(border)
        }

      {:ok, id} ->
        id
    end
  end

  def id_sorted_borders(db), do: DBUtil.id_sorted_values(db.borders)
end
