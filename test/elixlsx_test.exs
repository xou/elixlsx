defmodule ElixlsxTest do
  require Record
  Record.defrecord :xmlAttribute, Record.extract(:xmlAttribute, from_lib: "xmerl/include/xmerl.hrl")
  Record.defrecord :xmlText, Record.extract(:xmlText, from_lib: "xmerl/include/xmerl.hrl")

  use ExUnit.Case
  doctest Elixlsx
  doctest Elixlsx.Util
  doctest Elixlsx.XMLTemplates

  alias Elixlsx.XMLTemplates

  test "the truth" do
    assert 1 + 1 == 2
  end

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

end
