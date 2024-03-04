defmodule Elixlsx.Image do
  alias Elixlsx.Image

  @moduledoc ~S"""
  Structure for excel drawing files.
  - x_from_offset: integer
  - x_to_offset: integer
  - y_from_offset: integer
  - y_to_offset: integer
  - positioning: atom (:absolute, :oneCell, :twoCell)
  - width: integer
  - height: integer
  """

  defstruct file_path: "",
            type: "image/png",
            extension: "png",
            rowidx: 0,
            colidx: 0,
            x_from_offset: 0,
            y_from_offset: 0,
            x_to_offset: 0,
            y_to_offset: 0,
            positioning: "",
            width: 1,
            height: 1

  @type t :: %Image{
          file_path: String.t(),
          type: String.t(),
          extension: String.t(),
          rowidx: integer,
          colidx: integer,
          x_from_offset: integer,
          y_from_offset: integer,
          x_to_offset: integer,
          y_to_offset: integer,
          positioning: atom | String.t(),
          width: integer,
          height: integer
        }

  @doc """
  Create an image struct based on opts
  """
  def new(file_path, rowidx, colidx, opts \\ []) do
    {ext, type} = image_type(file_path)

    %Image{
      file_path: file_path,
      type: type,
      extension: ext,
      rowidx: rowidx,
      colidx: colidx,
      x_from_offset: Keyword.get(opts, :x_from_offset, 0),
      y_from_offset: Keyword.get(opts, :y_from_offset, 0),
      x_to_offset: Keyword.get(opts, :x_to_offset, 0),
      y_to_offset: Keyword.get(opts, :y_to_offset, 0),
      positioning: Keyword.get(opts, :positioning, ""),
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
