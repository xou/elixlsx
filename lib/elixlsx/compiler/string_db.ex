defmodule Elixlsx.Compiler.StringDB do
  alias Elixlsx.Compiler.StringDB
  @moduledoc ~S"""
  Strings in XLSX can be stored in a sharedStrings.xml file and be looked up
  by ID. This module handles collection of the data in the preprocessing phase.
  """
  defstruct strings: %{}, element_count: 0

  @type t :: %StringDB {
    strings: %{String.t => non_neg_integer},
    element_count: non_neg_integer
  }

  @spec register_string(StringDB.t, String.t) :: StringDB.t
  def register_string(stringdb, s) do
    case Map.fetch(stringdb.strings, s) do
      :error -> %StringDB{strings: Map.put(stringdb.strings, s, stringdb.element_count),
                          element_count: stringdb.element_count + 1}
      {:ok, _} -> stringdb
    end
  end

  def get_id(stringdb, s) do
    case Map.fetch(stringdb.strings, s) do
      :error ->
        raise %ArgumentError{
          message: "Invalid key provided for StringDB.get_id: " <> inspect(s)}
      {:ok, id} ->
        id
    end
  end

  def sorted_id_string_tuples(stringdb) do
    Enum.map(stringdb.strings, fn ({k, v}) -> {v, k} end) |> Enum.sort
  end
end
