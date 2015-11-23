
defmodule Elixlsx.Compiler.StringDB do
  alias Elixlsx.Compiler.StringDB
  @moduledoc ~S"""
  Strings in XLSX can be stored in a sharedStrings.xml file and be looked up
  by ID. This module handles collection of the data in the preprocessing phase.
  """
  defstruct strings: %{}, element_count: 0

  @type t :: %StringDB {
    strings: %{String.t => non_neg_integer},
    element_count: non_neg_integer
  }

  @spec register_string(StringDB.t, String.t) :: StringDB.t
  def register_string(stringdb, s) do
    case Dict.fetch(stringdb.strings, s) do
      :error -> %StringDB{strings: Dict.put(stringdb.strings, s, stringdb.element_count),
                          element_count: stringdb.element_count + 1}
      {:ok, _} -> stringdb
    end
  end

  def get_id(stringdb, s) do
    case Dict.fetch(stringdb.strings, s) do
      :error ->
        raise %ArgumentError{
          message: "Invalid key provided for StringDB.get_id: " <> inspect(s)}
      {:ok, id} ->
        id
    end
  end

  def sorted_id_string_tuples(stringdb) do
    Enum.map(stringdb.strings, fn ({k, v}) -> {v, k} end) |> Enum.sort
  end
end

defmodule Font do
  defstruct bold: false, italic: false, underline: false,
            strike: false, size: nil

  def from_props props do
    %Font{bold: !!props[:bold],
          italic: !!props[:italic],
          underline: !!props[:underline],
          strike: !!props[:strike],
          size: props[:size]
         }
  end


end


defmodule Elixlsx.Compiler.FontDB do
  alias Elixlsx.Compiler.FontDB
  defstruct fonts: %{}, element_count: 0

  @type t :: %FontDB {
    fonts: %{Font.t => pos_integer},
    element_count: non_neg_integer
  }

  @spec register_font(FontDB.t, Font.t) :: FontDB.t
  def register_font(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error -> %FontDB{fonts: Dict.put(fontdb.fonts, font, fontdb.element_count + 1),
                       element_count: fontdb.element_count + 1}
      {:ok, _} -> fontdb
    end
  end

  def get_id(fontdb, font) do
    case Dict.fetch(fontdb.fonts, font) do
      :error ->
        raise %ArgumentError{message: "Invalid key provided for FontDB.get_id: " <> inspect(font)}
      {:ok, id} ->
        id
    end
  end

  def id_sorted_fonts(fontdb) do
    fontdb.fonts
    |> Enum.map(fn ({k, v}) -> {v, k} end)
    |> Enum.sort
    |> Dict.values
  end
end


defmodule CellStyle do
  defstruct font: nil

  @type t :: %CellStyle{
    font: Font.t
  }


  def from_props props do
    font = Font.from_props props
    %CellStyle{font: font}
  end
end


defmodule Elixlsx.Compiler.CellStyleDB do
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.FontDB

  defstruct cellstyles: %{}, element_count: 0

  @type t :: %CellStyleDB {
    cellstyles: %{CellStyle.t => non_neg_integer},
    element_count: non_neg_integer
  }


  def register_style(cellstyledb, style) do
    case Dict.fetch(cellstyledb.cellstyles, style) do
      :error ->
        # add +1 here already since "0" refers to the default style
        csdb = update_in cellstyledb.cellstyles,
                  &(Dict.put &1, style, (cellstyledb.element_count + 1))
        update_in csdb.element_count, &(&1 + 1)
      {:ok, _} ->
        cellstyledb
    end
  end

  def get_id(cellstyledb, style) do
    case Dict.fetch(cellstyledb.cellstyles, style) do
      :error ->
        raise %ArgumentError{message: "Could not find key in styledb: " <> inspect(style)}
      {:ok, key} ->
        key
    end
  end

  def id_sorted_styles(cellstyledb) do
    cellstyledb.cellstyles
    |> Enum.map(fn ({k, v}) -> {v, k} end)
    |> Enum.sort
    |> Dict.values
  end

  @doc ~S"""
  Recursively register all fonts, border* and fill* properties (*=TBD)
  in the WorkbookCompInfo structure.
  """
  @spec register_all(WorkbookCompInfo.t) :: WorkbookCompInfo.t
  def register_all(wci) do
    Enum.reduce wci.cellstyledb.cellstyles, wci, fn ({style, _}, wci) ->
      update_in wci.fontdb, &(FontDB.register_font &1, style.font)
      # TODO: update_in wci.borderstyledb ...; wci.fillstyledb...
    end
  end
end


defmodule Elixlsx.Compiler.SheetCompInfo do
  @moduledoc ~S"""
  Compilation info for a sheet, to be filled during the actual
  write process.
  """
  defstruct rId: "", filename: "sheet1.xml", sheetId: 0
  @type t :: %Elixlsx.Compiler.SheetCompInfo{
    rId: String.t,
    filename: String.t,
    sheetId: non_neg_integer
  }

  @spec make(non_neg_integer, non_neg_integer) :: SheetCompInfo.t
  def make sheetidx, rId do
    %Elixlsx.Compiler.SheetCompInfo{rId: "rId" <> to_string(rId),
                   filename: "sheet" <> to_string(sheetidx) <> ".xml",
                   sheetId: sheetidx}
  end
end


defmodule Elixlsx.Compiler.WorkbookCompInfo do
  @moduledoc ~S"""
  This module aggregates information about the metainformation
  required to generate the XML file.

  It is used as the aggregator when folding over the individual
  cells.
  """
  defstruct sheet_info: nil,
            stringdb: %Elixlsx.Compiler.StringDB{},
            fontdb: %Elixlsx.Compiler.FontDB{},
            cellstyledb: %Elixlsx.Compiler.CellStyleDB{},
            next_free_xl_rid: nil
end


defmodule Elixlsx.Compiler do
  alias Elixlsx.Compiler.WorkbookCompInfo
  alias Elixlsx.Compiler.SheetCompInfo
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.StringDB

  @doc ~S"""
  Accepts a list of Sheets and the next free relationship ID.
  Returns a tuple containing a list of SheetCompInfo's and the next free
  relationship ID.
  """
  @spec make_sheet_info(nonempty_list(Sheet.t), non_neg_integer) :: {list(SheetCompInfo.t), non_neg_integer}
  def make_sheet_info(sheets, init_rId) do
    # fold helper. aggregator holds {list(SheetCompInfo), sheetidx, rId}.
    add_sheet =
      fn (_, {sci, idx, rId}) ->
        {[SheetCompInfo.make(idx, rId) | sci], idx + 1, rId + 1}
      end

    # TODO probably better to use a zip [1..] |> map instead of fold[l|r]/reverse
    {sheetCompInfos, _, nextrID} = List.foldl(sheets, {[], 1, init_rId}, add_sheet)
    {Enum.reverse(sheetCompInfos), nextrID}
  end

  def compinfo_cell_pass_value wci, value do
    cond do
      is_binary(value) && String.valid?(value)
        -> update_in wci.stringdb, &StringDB.register_string(&1, value)
      true -> wci
    end
  end


  def compinfo_cell_pass_style wci, props do
    update_in wci.cellstyledb, &CellStyleDB.register_style(&1, CellStyle.from_props(props))
  end


  @spec compinfo_cell_pass(WorkbookCompInfo.t, any) :: WorkbookCompInfo.t
  def compinfo_cell_pass wci, cell do
    cond do
      is_list(cell) ->
        wci
        |> compinfo_cell_pass_value(hd cell)
        |> compinfo_cell_pass_style(tl cell)
      true ->
        compinfo_cell_pass_value wci, cell
    end
  end


  @spec compinfo_from_rows(WorkbookCompInfo.t, list(list(any()))) :: WorkbookCompInfo.t
  def compinfo_from_rows wci, rows do
    List.foldl rows, wci, fn (cols, wci) ->
      List.foldl cols, wci, fn (cell, wci) ->
        compinfo_cell_pass wci, cell
      end
    end
  end

  @spec compinfo_from_sheets(WorkbookCompInfo.t, list(Sheet.t)) :: WorkbookCompInfo.t
  def compinfo_from_sheets wci, sheets do
    List.foldl sheets, wci, fn (sheet, wci) ->
      compinfo_from_rows wci, sheet.rows
    end
  end

  @first_free_rid 2
  def make_workbook_comp_info workbook do
    {sci, next_rId} = make_sheet_info(workbook.sheets, @first_free_rid)

    %WorkbookCompInfo{
      sheet_info: sci,
      next_free_xl_rid: next_rId,
    }
    |> compinfo_from_sheets(workbook.sheets)
    |> CellStyleDB.register_all
  end
end
