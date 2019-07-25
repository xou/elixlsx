defmodule Elixlsx.Style.Protection do
  @moduledoc ~S"""
  Protection properties.

  Supported formatting properties are:

  - locked: boolean
  """
  alias __MODULE__

  defstruct locked: nil

  @type t :: %Protection{
          locked: boolean
        }

  @doc ~S"""
  Create a Protection object from a property list.
  """
  def from_props(props) do
    ft = %Protection{locked: props[:locked]}

    if ft == %Protection{}, do: nil, else: ft
  end
end
