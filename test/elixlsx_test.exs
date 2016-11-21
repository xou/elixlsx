defmodule ElixlsxTest do
  require Record
  Record.defrecord :xmlAttribute, Record.extract(:xmlAttribute, from_lib: "xmerl/include/xmerl.hrl")
  Record.defrecord :xmlText, Record.extract(:xmlText, from_lib: "xmerl/include/xmerl.hrl")

  use ExUnit.Case
  doctest Elixlsx
  doctest Elixlsx.Sheet
  doctest Elixlsx.Util, import: true
  doctest Elixlsx.XMLTemplates

  alias Elixlsx.XMLTemplates
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.Style.Font

  def xpath(el, path) do
    :xmerl_xpath.string(to_char_list(path), el)
  end

  defp to_text xml_text do
    xmlText(value: value) = xml_text
    to_string value
  end

  test "basic StringDB functionality" do
    sdb = (%StringDB{}
            |> StringDB.register_string("Hello")
            |> StringDB.register_string("World")
            |> StringDB.register_string("Hello"))

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples sdb)

    {xmerl, []} = :xmerl_scan.string String.to_char_list(xml)

    strings = :xmerl_xpath.string('/sst/si/t/text()', xmerl)

    assert length(strings) == 2
    [sis1, sis2] = strings

    assert to_text(sis1) == "Hello"
    assert to_text(sis2) == "World"
  end

  test "xml escaping StringDB functionality" do
    sdb = (%StringDB{}
            # An unfortunate side effect of :xmerl_scan is that although some values
            # will be xml escaped, for example "&" replaced with "&amp;", the parser
            # still splits on the "&" of the escaped value, thus creating two values
            # instead of one. This will not effect the actual output of the library
            # though.
            |> StringDB.register_string("Hello World & Goodbye Cruel World"))

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples sdb)

    {xmerl, []} = :xmerl_scan.string String.to_char_list(xml)

    strings = :xmerl_xpath.string('/sst/si/t/text()', xmerl)

    assert length(strings) == 2
    [sis1, sis2] = strings

    assert to_text(sis1) <> to_text(sis2) == "Hello World & Goodbye Cruel World"
  end

  test "font color" do
    xml = Font.from_props(color: "#012345") |>
    Font.get_stylexml_entry

    {xmerl, []} = :xmerl_scan.string String.to_char_list(xml)

    [color] = :xmerl_xpath.string('/font/color/@rgb', xmerl)

    assert xmlAttribute(color, :value) == 'FF012345'
  end

  test "font name" do
    xml = Font.from_props(name: "Arial") |>
    Font.get_stylexml_entry

    {xmerl, []} = :xmerl_scan.string String.to_char_list(xml)

    [name] = :xmerl_xpath.string('/font/name/@val', xmerl)

    assert xmlAttribute(name, :value) == 'Arial'
  end
end
