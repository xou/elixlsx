defmodule Elixlsx.Compiler.FontDB do
  alias __MODULE__
  alias Elixlsx.Style.Font

  defstruct fonts: %{}, element_count: 0

  @type t :: %FontDB {
    fonts: %{Font.t => pos_integer},
    element_count: non_neg_integer
  }

  @spec register_font(FontDB.t, Font.t) :: FontDB.t
  def register_font(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error -> %FontDB{fonts: Dict.put(fontdb.fonts, font, fontdb.element_count + 1),
                       element_count: fontdb.element_count + 1}
      {:ok, _} -> fontdb
    end
  end

  def get_id(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for FontDB.get_id: " <> inspect(font)}
      {:ok, id} ->
        id
    end
  end

  def id_sorted_fonts(fontdb) do
    fontdb.fonts
    |> Enum.map(fn ({k, v}) -> {v, k} end)
    |> Enum.sort
    |> Dict.values
  end
end
