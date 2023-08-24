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
            height: 1,
            binary: nil,
            align_x: :left

  @type t :: %Image{
          file_path: String.t() | {String.t(), binary},
          type: String.t(),
          extension: String.t(),
          x: integer,
          y: integer,
          x_offset: integer,
          y_offset: integer,
          width: integer,
          height: integer,
          binary: binary | nil,
          align_x: :left | :right
        }

  @doc """
  Create an image struct based on opts
  """
  def new(file_data, x, y, opts \\ []) do
    {path, {ext, type}, binary} =
      case file_data do
        {path, binary} -> {path, image_type(path), binary}
        path -> {path, image_type(path), nil}
      end

    %Image{
      file_path: path,
      binary: binary,
      type: type,
      extension: ext,
      x: x,
      y: y,
      x_offset: Keyword.get(opts, :x_offset, 0),
      y_offset: Keyword.get(opts, :y_offset, 0),
      width: Keyword.get(opts, :width, 1),
      height: Keyword.get(opts, :height, 1),
      align_x: Keyword.get(opts, :align_x, :left)
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
