defmodule Elixlsx.Image do
  alias Elixlsx.Image

  @moduledoc ~S"""
  An Image can either by a path to an image, or
  a binary {"path_or_unique_id", <<binary>>}

  When aligning the image to the right you might
  need to adjust the char attribute. char is the
  max character width of a font, this is used when
  calculating how many pixels are in a column.
  You might need to experiment with this value
  depending on what font and size you are using.
  """

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
            align_x: :left,
            char: 7

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
          align_x: :left | :right,
          char: integer
        }

  @doc """
  Create an image struct based on opts
  """
  def new(_, _, _, opts \\ [])

  def new(file_path, x, y, opts) when is_binary(file_path) do
    new({file_path, nil}, x, y, opts)
  end

  def new({file_path, binary}, x, y, opts) do
    {ext, type} = image_type(file_path)

    %Image{
      file_path: file_path,
      binary: binary,
      type: type,
      extension: ext,
      x: x,
      y: y,
      x_offset: Keyword.get(opts, :x_offset, 0),
      y_offset: Keyword.get(opts, :y_offset, 0),
      width: Keyword.get(opts, :width, 1),
      height: Keyword.get(opts, :height, 1),
      align_x: Keyword.get(opts, :align_x, :left),
      char: Keyword.get(opts, :char, 7)
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
