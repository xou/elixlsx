defmodule Elixlsx.Compiler.NumFmtDB do
  alias __MODULE__
  alias Elixlsx.Style.NumFmt
  alias Elixlsx.Compiler.DBUtil
  defstruct numfmts: %{}, nextid: 164

  @type t :: %NumFmtDB{
          numfmts: %{NumFmt.t() => pos_integer},
          nextid: non_neg_integer
        }

  def register_numfmt(db, value) do
    {dict, ec} = DBUtil.register({db.numfmts, db.nextid}, value)
    %NumFmtDB{numfmts: dict, nextid: ec}
  end

  @doc ~S"""
  register an ID for a built-in NumFmt object.

  built-in refers to the 164 objects (ids 0-163) that are
  defined or reserved in the XLSX standard. A NumFmt object
  mimicking the behaviour of such a built-in style can be
  associated with the built-in id using this function, which
  should save a couple of bytes in the resulting XLSX file.
  """
  def register_builtin(db, value, id) do
    update_in(db.numfmts, &Map.put(&1, value, id))
  end

  def get_id(db, value), do: DBUtil.get_id(db.numfmts, value)

  @spec id_sorted_numfmts(NumFmtDB.t()) :: list(NumFmt.t())
  def id_sorted_numfmts(db), do: DBUtil.id_sorted_values(db.numfmts)

  @doc ~S"""
  Return a list of tuples {id, NumFmt.t} for all custom (id >= 164)
  NumFmts.
  """
  @spec custom_numfmt_id_tuples(NumFmtDB.t()) :: list({non_neg_integer, NumFmt.t()})
  def custom_numfmt_id_tuples(db) do
    db.numfmts
    |> Enum.map(fn {k, v} -> {v, k} end)
    |> Enum.sort()
    |> Enum.filter(fn {id, _} -> id >= 164 end)
  end
end
