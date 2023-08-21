defmodule Elixlsx.Compiler.LinkDB do
  alias __MODULE__

  defstruct links: %{}, element_count: 0

  @type t :: %LinkDB{
          links: %{String.t() => String.t()},
          element_count: non_neg_integer
        }

  @spec register_link(LinkDB.t(), String.t(), non_neg_integer) :: LinkDB.t()
  def register_link(linkdb, link, rId) do
    case Map.fetch(linkdb.links, link) do
      :error ->
        %LinkDB{
          links: Map.put(linkdb.links, link, "rId" <> to_string(rId)),
          element_count: linkdb.element_count + 1
        }

      {:ok, _} ->
        linkdb
    end
  end

  def get_id(linkdb, link) do
    case Map.fetch(linkdb.links, link) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for LinkDB.get_id: " <> inspect(link)}

      {:ok, id} ->
        id
    end
  end
end
