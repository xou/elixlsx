defmodule Elixlsx.Image do
  alias Elixlsx.Image

  defstruct file_path: "",
            type: "image/png",
            extension: "png",
            x: 0,
            y: 0,
            x_offset: 0,
            y_offset: 0,
            width: 1,
            height: 1

  @type t :: %Image{
          file_path: String.t(),
          type: String.t(),
          extension: String.t(),
          x: integer,
          y: integer,
          x_offset: integer,
          y_offset: integer,
          width: integer,
          height: integer
        }

  @doc """
  Create an image struct based on opts
  """
  def new(file_path, x, y, opts \\ []) do
    {ext, type} = image_type(file_path)

    %Image{
      file_path: file_path,
      type: type,
      extension: ext,
      x: x,
      y: y,
      x_offset: Keyword.get(opts, :x_offset, 0),
      y_offset: Keyword.get(opts, :y_offset, 0),
      width: Keyword.get(opts, :width, 1),
      height: Keyword.get(opts, :height, 1)
    }
  end

  defp image_type(file_path) do
    case Path.extname(file_path) do
      ".jpg" -> {"jpg", "image/jpeg"}
      ".jpeg" -> {"jpeg", "image/jpeg"}
      ".png" -> {"png", "image/png"}
    end
  end
end
