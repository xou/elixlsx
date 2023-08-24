defmodule Elixlsx.Compiler.DrawingDB do
  alias __MODULE__
  alias Elixlsx.Compiler.DBUtil
  alias Elixlsx.Image

  @doc """
  Database of drawing elements in the whole document. drawing id values must be
  unique across the document regardless of what kind of drawing they are.
  So far this only supports images, but could be extended to include other
  kinds of drawing.
  An alternative would be to add a Drawing module and have "subclasses" for
  different drawing types
  """

  defstruct images: %{}, element_count: 0

  @type t :: %DrawingDB{
          images: %{Image.t() => pos_integer},
          element_count: non_neg_integer
        }

  def register_image(drawingdb, image) do
    case Map.fetch(drawingdb.images, image) do
      :error ->
        %DrawingDB{
          images: Map.put(drawingdb.images, image, drawingdb.element_count + 1),
          element_count: drawingdb.element_count + 1
        }

      {:ok, _} ->
        drawingdb
    end
  end

  def get_id(drawingdb, image) do
    case Map.fetch(drawingdb.images, image) do
      :error ->
        raise %ArgumentError{
          message: "Invalid key provided for DrawingDB.get_id: " <> inspect(image)
        }

      {:ok, id} ->
        id
    end
  end

  def id_sorted_drawings(db), do: DBUtil.id_sorted_values(db.images)

  def image_types(db) do
    db.images
    |> Enum.reduce(%MapSet{}, fn {i, _}, acc ->
      MapSet.put(acc, {i.extension, i.type})
    end)
    |> Enum.to_list()
  end
end
