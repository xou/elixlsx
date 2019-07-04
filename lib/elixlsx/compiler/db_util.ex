defmodule Elixlsx.Compiler.DBUtil do
  @moduledoc ~S"""
  Generic functions for the Compiler.*DB modules.
  """
  @type object_type :: any
  @type gen_db_datatype :: %{object_type => non_neg_integer}
  @type gen_db_type :: {gen_db_datatype, non_neg_integer}

  @doc ~S"""
  If the value does not exist in the database, return
  the tuple {dict, nextid} unmodified. Otherwise,
  returns a tuple {dict', nextid+1}, where dict'
  is the dictionary with the new element inserted
  (with id `nextid`)
  """
  @spec register(gen_db_type, object_type) :: gen_db_type
  def register({dict, nextid}, value) do
    # Note that the parameter "value" in the API
    # refers to the *key* in the dictionary
    case Map.fetch(dict, value) do
      :error -> {Map.put(dict, value, nextid), nextid + 1}
      {:ok, _} -> {dict, nextid}
    end
  end

  @doc ~S"""
  return the ID for an object in the database
  """
  @spec get_id(gen_db_datatype, object_type) :: non_neg_integer
  def get_id(dict, value) do
    case Map.fetch(dict, value) do
      :error -> raise %ArgumentError{message: "Unable to find element: " <> inspect(value)}
      {:ok, id} -> id
    end
  end

  @spec id_sorted_values(gen_db_datatype) :: list(object_type)
  def id_sorted_values(dict) do
    dict
    |> Enum.map(fn {k, v} -> {v, k} end)
    |> Enum.sort()
    |> Enum.map(fn {_, k} -> k end)
  end
end
