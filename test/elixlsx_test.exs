defmodule ElixlsxTest do
  alias Elixlsx.Sheet
  require Record

  Record.defrecord(
    :xmlAttribute,
    Record.extract(:xmlAttribute, from_lib: "xmerl/include/xmerl.hrl")
  )

  Record.defrecord(:xmlElement, Record.extract(:xmlElement, from_lib: "xmerl/include/xmerl.hrl"))
  Record.defrecord(:xmlText, Record.extract(:xmlText, from_lib: "xmerl/include/xmerl.hrl"))

  use ExUnit.Case
  doctest Elixlsx
  doctest Elixlsx.Sheet
  doctest Elixlsx.XMLTemplates
  doctest Elixlsx.Color
  doctest Elixlsx.Style.Border
  doctest Elixlsx.Style.BorderStyle

  alias Elixlsx.XMLTemplates
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.Style.Font
  alias Elixlsx.Workbook
  alias Elixlsx.Sheet

  def xpath(el, path) do
    :xmerl_xpath.string(to_charlist(path), el)
  end

  defp xml_inner_strings(xml, path) do
    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    Enum.map(
      xpath(xmerl, path),
      fn element ->
        Enum.reduce(xmlElement(element, :content), "", fn text, acc ->
          acc <> to_text(text)
        end)
      end
    )
  end

  defp to_text(xml_text) do
    xmlText(value: value) = xml_text
    to_string(value)
  end

  test "basic StringDB functionality" do
    sdb =
      %StringDB{}
      |> StringDB.register_string("Hello")
      |> StringDB.register_string("World")
      |> StringDB.register_string("Hello")

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(sdb))

    assert xml_inner_strings(xml, '/sst/si/t') == ["Hello", "World"]
  end

  test "xml escaping StringDB functionality" do
    sdb =
      %StringDB{}
      |> StringDB.register_string("Hello World & Goodbye Cruel World")

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(sdb))

    assert xml_inner_strings(xml, '/sst/si/t') == ["Hello World & Goodbye Cruel World"]
  end

  test "font color" do
    xml =
      Font.from_props(color: "#012345")
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [color] = :xmerl_xpath.string('/font/color/@rgb', xmerl)

    assert xmlAttribute(color, :value) == 'FF012345'
  end

  test "font name" do
    xml =
      Font.from_props(font: "Arial")
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [name] = :xmerl_xpath.string('/font/name/@val', xmerl)

    assert xmlAttribute(name, :value) == 'Arial'
  end

  test "blank sheet name" do
    sheet1 = Sheet.with_name("")

    assert_raise ArgumentError, "The sheet name cannot be blank.", fn ->
      %Workbook{sheets: [sheet1]}
      |> Elixlsx.write_to("test.xlsx")
    end
  end

  test "too long sheet name" do
    sheet1 = Sheet.with_name("This is a very looong sheet name")

    assert_raise ArgumentError,
                 ~r/The sheet name .* is too long. Maximum 31 chars allowed for name./,
                 fn ->
                   %Workbook{sheets: [sheet1]}
                   |> Elixlsx.write_to("test.xlsx")
                 end
  end

  test "invalid chars in sheet name" do
    sheet_names = ~W(: \ / ? * [ ])

    Enum.each(sheet_names, fn name ->
      sheet1 = Sheet.with_name(name)

      assert_raise ArgumentError,
                   ~r/The sheet name .* contains following invalid characters:/,
                   fn ->
                     %Workbook{sheets: [sheet1]}
                     |> Elixlsx.write_to("test.xlsx")
                   end
    end)
  end

  test "sheet rels" do
    workbook = %Workbook{
      sheets: [
        Sheet.with_name("foo")
        |> Sheet.insert_image(0, 0, "ladybug-3475779_640.jpg"),
        Sheet.with_name("bar")
        |> Sheet.insert_image(0, 0, "ladybug-3475779_640.jpg")
      ]
    }

    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)

    doc = get_doc(res, 'xl/worksheets/sheet1.xml')
    assert Floki.find(doc, "drawing") == [{"drawing", [{"r:id", "rId1"}], []}]

    doc = get_doc(res, 'xl/worksheets/sheet2.xml')
    assert Floki.find(doc, "drawing") == [{"drawing", [{"r:id", "rId1"}], []}]

    doc = get_doc(res, 'xl/worksheets/_rels/sheet1.xml.rels')
    rel = Floki.find(doc, "relationship")
    assert Floki.attribute(rel, "id") == ["rId1"]
    assert Floki.attribute(rel, "target") == ["../drawings/drawing1.xml"]

    # target should be drawing 2, but the id
    # should be the same as sheet 1
    doc = get_doc(res, 'xl/worksheets/_rels/sheet2.xml.rels')
    rel = Floki.find(doc, "relationship")
    assert Floki.attribute(rel, "id") == ["rId1"]
    assert Floki.attribute(rel, "target") == ["../drawings/drawing2.xml"]
  end

  test "drawing single cell" do
    workbook = %Workbook{
      sheets: [
        %Sheet{name: "single", emu: 10}
        |> Sheet.set_col_width("A", "100px")
        |> Sheet.set_row_height(1, "100px")
        |> Sheet.insert_image(0, 0, "ladybug-3475779_640.jpg", width: 100, height: 100)
      ]
    }

    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)

    doc = get_doc(res, 'xl/drawings/drawing1.xml')

    assert Floki.find(doc, "from") == [
             {"xdr:from", [],
              [
                {"xdr:col", [], ["0"]},
                {"xdr:coloff", [], ["0"]},
                {"xdr:row", [], ["0"]},
                {"xdr:rowoff", [], ["0"]}
              ]}
           ]

    assert Floki.find(doc, "to") == [
             {"xdr:to", [],
              [
                {"xdr:col", [], ["0"]},
                {"xdr:coloff", [], ["1000"]},
                {"xdr:row", [], ["0"]},
                {"xdr:rowoff", [], ["1000"]}
              ]}
           ]

    assert Floki.find(doc, "xfrm") == [
             {"a:xfrm", [],
              [
                {"a:off", [{"x", "0"}, {"y", "0"}], []},
                {"a:ext", [{"cx", "1000"}, {"cy", "1000"}], []}
              ]}
           ]
  end

  test "drawing multi cell" do
    workbook = %Workbook{
      sheets: [
        %Sheet{name: "multi", emu: 10}
        |> Sheet.set_col_width("A", "100px")
        |> Sheet.set_col_width("B", "100px")
        |> Sheet.set_col_width("C", "100px")
        |> Sheet.set_row_height(1, "100px")
        |> Sheet.set_row_height(2, "100px")
        |> Sheet.set_row_height(3, "100px")
        |> Sheet.insert_image(1, 1, "ladybug-3475779_640.jpg", width: 200, height: 200)
      ]
    }

    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)

    doc = get_doc(res, 'xl/drawings/drawing1.xml')

    assert Floki.find(doc, "from") == [
             {"xdr:from", [],
              [
                {"xdr:col", [], ["1"]},
                {"xdr:coloff", [], ["0"]},
                {"xdr:row", [], ["1"]},
                {"xdr:rowoff", [], ["0"]}
              ]}
           ]

    assert Floki.find(doc, "to") == [
             {"xdr:to", [],
              [
                {"xdr:col", [], ["2"]},
                {"xdr:coloff", [], ["1000"]},
                {"xdr:row", [], ["2"]},
                {"xdr:rowoff", [], ["1000"]}
              ]}
           ]

    assert Floki.find(doc, "xfrm") == [
             {"a:xfrm", [],
              [
                {"a:off", [{"x", "1000"}, {"y", "1000"}], []},
                {"a:ext", [{"cx", "2950"}, {"cy", "3000"}], []}
              ]}
           ]
  end

  test "set_col_width" do
    sheet =
      %Sheet{}
      |> Sheet.set_col_width("A", "131px")
      |> Sheet.set_col_width("B", 20)

    assert sheet.col_widths == %{
             1 => 18.0,
             2 => 20
           }
  end

  test "set_col_width max_char_width" do
    sheet =
      %Sheet{max_char_width: 8}
      |> Sheet.set_col_width("A", "131px")
      |> Sheet.set_col_width("B", 20)

    assert sheet.col_widths == %{
             1 => 15.75,
             2 => 20
           }
  end

  test "set_row_height" do
    sheet =
      %Sheet{}
      |> Sheet.set_row_height(1, "20px")
      |> Sheet.set_row_height(2, 15)

    assert sheet.row_heights == %{
             1 => 15,
             2 => 15
           }
  end

  test "set_max_char_width" do
    sheet =
      %Sheet{}
      |> Sheet.set_max_char_width(10)

    assert sheet.max_char_width == 10
  end

  defp get_doc(res, name) do
    {_, sheet} = Enum.find(res, fn {a, _} -> a == name end)
    Floki.parse_fragment!(sheet)
  end
end
