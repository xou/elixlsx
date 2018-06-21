defmodule Elixlsx.Image do
  alias Elixlsx.Image

  @moduledoc ~S"""
  Structure for excel drawing files.

  - x_offset: integer
  - y_offset: integer
  - x_scale: float
  - y_scale: float
  - positioning: atom (:absolute, :oneCell, :twoCell)
  """

  defstruct file_path: "",
            type: "image/png",
            extension: "png",
            rowidx: 0,
            colidx: 0,
            x_offset: 0,
            y_offset: 0,
            x_scale: 1,
            y_scale: 1,
            positioning: :twoCell

  @type t :: %Image{
          file_path: String.t(),
          type: String.t(),
          extension: String.t(),
          rowidx: integer,
          colidx: integer,
          x_offset: integer,
          y_offset: integer,
          x_scale: float,
          y_scale: float,
          positioning: atom
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
      x_offset: Keyword.get(opts, :x_offset, 0),
      y_offset: Keyword.get(opts, :y_offset, 0),
      x_scale: Keyword.get(opts, :x_scale, 1),
      y_scale: Keyword.get(opts, :y_scale, 1),
      positioning: Keyword.get(opts, :positioning, :twoCell)
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
