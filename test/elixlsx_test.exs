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

    assert xml_inner_strings(xml, ~c"/sst/si/t") == ["Hello", "World"]
  end

  test "xml escaping StringDB functionality" do
    sdb =
      %StringDB{}
      |> StringDB.register_string("Hello World & Goodbye Cruel World")

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(sdb))

    assert xml_inner_strings(xml, ~c"/sst/si/t") == ["Hello World & Goodbye Cruel World"]
  end

  test "font color" do
    xml =
      Font.from_props(color: "#012345")
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [color] = :xmerl_xpath.string(~c"/font/color/@rgb", xmerl)

    assert xmlAttribute(color, :value) == ~c"FF012345"
  end

  test "font name" do
    xml =
      Font.from_props(font: "Arial")
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [name] = :xmerl_xpath.string(~c"/font/name/@val", xmerl)

    assert xmlAttribute(name, :value) == ~c"Arial"
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

  test "docProps/app" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"docProps/app.xml")

    expected = """
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Properties
        xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
        xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        <TotalTime>0</TotalTime>
        <Application>Elixlsx</Application>
        <AppVersion>0.52</AppVersion>
      </Properties>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "docProps/core" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"docProps/core.xml")

    assert [
             {:pi, "xml", [{"version", "1.0"}, {"encoding", "UTF-8"}, {"standalone", "yes"}]},
             {
               "cp:coreproperties",
               [
                 {"xmlns:cp",
                  "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"},
                 {"xmlns:dc", "http://purl.org/dc/elements/1.1/"},
                 {"xmlns:dcterms", "http://purl.org/dc/terms/"},
                 {"xmlns:dcmitype", "http://purl.org/dc/dcmitype/"},
                 {"xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"}
               ],
               [
                 {"dcterms:created", [{"xsi:type", "dcterms:W3CDTF"}], [_]},
                 {"dc:language", [], ["en-US"]},
                 {"dcterms:modified", [{"xsi:type", "dcterms:W3CDTF"}], [_]},
                 {"cp:revision", [], ["1"]}
               ]
             }
           ] = doc
  end

  test "_rels/.rels" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"_rels/.rels")

    expected = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship
        Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
        Target="xl/workbook.xml"
      />
      <Relationship
        Id="rId2"
        Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
        Target="docProps/core.xml"
      />
      <Relationship
        Id="rId3"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
        Target="docProps/app.xml"
      />
    </Relationships>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "xl/styles.xml" do
    workbook = %Workbook{
      sheets: [Sheet.with_name("foo")],
      font: "Calibri Light",
      font_size: 16
    }

    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"xl/styles.xml")

    expected = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <fonts count="1">
        <font>
          <name val="Calibri Light" />
          <sz val="16" />
        </font>
      </fonts>
      <fills count="2">
        <fill>
          <patternFill patternType="none" />
        </fill>
        <fill>
          <patternFill patternType="gray125" />
        </fill>
      </fills>
      <borders count="1">
        <border />
      </borders>
      <cellStyleXfs count="1">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" applyAlignment="1">
          <alignment wrapText="1" />
        </xf>
      </cellStyleXfs>
      <cellXfs count="1">
        <xf borderId="0" numFmtId="0" fillId="0" fontId="0" xfId="0"/>
      </cellXfs>
    </styleSheet>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "xl/sharedStrings.xml" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"xl/sharedStrings.xml")

    expected = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <sst
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      count="0"
      uniqueCount="0">
    </sst>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "xl/workbook.xml" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"xl/workbook.xml")

    expected = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <workbook
      xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <fileVersion appName="Calc" />
      <bookViews>
        <workbookView activeTab="0" />
      </bookViews>
      <sheets>
        <sheet name="foo" sheetId="1" state="visible" r:id="rId2" />
      </sheets>
      <calcPr fullCalcOnLoad="1" iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001" />
    </workbook>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "xl/_rels/workbook.xml.rels" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"xl/_rels/workbook.xml.rels")

    expected = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship
        Id="rId1"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
        Target="styles.xml"
      />
      <Relationship
        Id="rId2"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
        Target="worksheets/sheet1.xml"
      />
      <Relationship
        Id="rId3"
        Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
        Target="sharedStrings.xml"
      />
    </Relationships>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "xl/worksheets/sheet1.xml" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"xl/worksheets/sheet1.xml")

    expected = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr filterMode="false">
        <pageSetUpPr fitToPage="false" />
      </sheetPr>
      <dimension ref="A1" />
      <sheetViews>
        <sheetView workbookViewId="0">
        <selection  activeCell="A1" sqref="A1" />
      </sheetView>
      </sheetViews>
      <sheetFormatPr defaultRowHeight="12.8" />
      <sheetData></sheetData>
      <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5" />
    </worksheet>
    """

    assert doc == Floki.parse_document!(expected)
  end

  test "[Content_Types].xml" do
    workbook = %Workbook{sheets: [Sheet.with_name("foo")]}
    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)
    doc = get_doc(res, ~c"[Content_Types].xml")

    expected = """
    <?xml version="1.0" encoding="UTF-8"?>
    <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
      <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
      <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
      <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
      <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
      <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
      <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
      <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
      <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
      <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
    </Types>
    """

    assert doc == Floki.parse_document!(expected)
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

    doc = get_doc(res, ~c"xl/worksheets/sheet1.xml")
    assert Floki.find(doc, "drawing") == [{"drawing", [{"r:id", "rId1"}], []}]

    doc = get_doc(res, ~c"xl/worksheets/sheet2.xml")
    assert Floki.find(doc, "drawing") == [{"drawing", [{"r:id", "rId1"}], []}]

    doc = get_doc(res, ~c"xl/worksheets/_rels/sheet1.xml.rels")
    rel = Floki.find(doc, "relationship")
    assert Floki.attribute(rel, "id") == ["rId1"]
    assert Floki.attribute(rel, "target") == ["../drawings/drawing1.xml"]

    # target should be drawing 2, but the id
    # should be the same as sheet 1
    doc = get_doc(res, ~c"xl/worksheets/_rels/sheet2.xml.rels")
    rel = Floki.find(doc, "relationship")
    assert Floki.attribute(rel, "id") == ["rId1"]
    assert Floki.attribute(rel, "target") == ["../drawings/drawing2.xml"]
  end

  test "drawing single cell" do
    workbook = %Workbook{
      sheets: [
        %Sheet{name: "single"}
        |> Sheet.set_col_width("A", 12)
        |> Sheet.set_row_height(1, 75)
        |> Sheet.insert_image(0, 0, "ladybug-3475779_640.jpg",
          width: 100,
          height: 100,
          char: 10,
          emu: 10
        )
      ]
    }

    wci = Elixlsx.Compiler.make_workbook_comp_info(workbook)
    res = Elixlsx.Writer.create_files(workbook, wci)

    doc = get_doc(res, ~c"xl/drawings/drawing1.xml")

    xml = """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <xdr:wsDr
      xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
      xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
      xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
      xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
      xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"
      xmlns:x3Unk="http://schemas.microsoft.com/office/drawing/2010/slicer"
      xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer">
      <xdr:oneCellAnchor>
        <xdr:from>
          <xdr:col>0</xdr:col>
          <xdr:colOff>0</xdr:colOff>
          <xdr:row>0</xdr:row>
          <xdr:rowOff>0</xdr:rowOff>
        </xdr:from>
        <xdr:ext cx="952500" cy="952500"/>
        <xdr:pic>
          <xdr:nvPicPr>
            <xdr:cNvPr id="0" name="image1.png" title="Image"/>
            <xdr:cNvPicPr preferRelativeResize="0"/>
          </xdr:nvPicPr>
          <xdr:blipFill>
            <a:blip cstate="print" r:embed="rId1"/>
            <a:stretch>
    	        <a:fillRect/>
            </a:stretch>
          </xdr:blipFill>
          <xdr:spPr>
            <a:prstGeom prst="rect">
              <a:avLst/>
            </a:prstGeom>
            <a:noFill/>
          </xdr:spPr>
        </xdr:pic>
        <xdr:clientData fLocksWithSheet="0"/>
      </xdr:oneCellAnchor>
    </xdr:wsDr>
    """

    assert doc == Floki.parse_document!(xml)
  end

  defp get_doc(res, name) do
    {_, sheet} = Enum.find(res, fn {a, _} -> a == name end)
    Floki.parse_fragment!(sheet)
  end
end
