defmodule Elixlsx.Compiler.DrawingDB do
  alias __MODULE__
  alias Elixlsx.Image

  @moduledoc ~S"""
  Database of drawing elements in the whole document.

  Drawing id values must be unique across the document
  regardless of what kind of drawing they are.
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

  def image_types(db) do
    db.images
    |> Enum.map(fn {i, _} -> {i.extension, i.type} end)
    |> Enum.uniq()
  end
end
