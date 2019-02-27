defmodule Elixlsx.SheetTest do
  use ExUnit.Case, async: false
  use ExCheck
  
  alias Elixlsx.Sheet
  
  property :sheet_name do
    for_all x in binary() do
      Sheet.with_name(x).name == x
    end
  end
  
  property :sheet_cols do
    for_all x in choose(65,90) do
      sheet = 
      %Sheet{}
      |> Sheet.set_col_width(<<x>>, 10)
      |> Sheet.set_col(<<x>>, bg_color: "#FFFF00", num_format: "mmm-yyyy")
      sheet.cols[x - 64] == %{bg_color: "#FFFF00", num_format: "mmm-yyyy", width: 10}
    end
  end
end
