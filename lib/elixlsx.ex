defmodule Sheet do
  defstruct name: "", rows: []
end

defmodule Elixlsx do
  
  @doc ~S"""
    returns a tuple {'docProps/app.xml', "XML Data"}
    TODO: This is stolen from libreoffice. Write reasonable stuff here.
  """
  def get_docProps_app_xml(data) do
    {'docProps/app.xml',
      ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Elixlsx/0.0.1$Linux_X86_64 nweh_project/00m0$Build-2</Application>
</Properties>
"""
	}
  end

  def get_docProps_core_xml(data) do
    {'docProps/core.xml',
      ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">2015-11-16T19:38:34Z</dcterms:created>
  <dc:language>en-US</dc:language>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2015-11-16T19:38:56Z</dcterms:modified>
  <cp:revision>1</cp:revision>
</cp:coreProperties>
"""
		}
  end


	def get_docProps_dir(data) do
		[get_docProps_app_xml(data), get_docProps_core_xml(data)]
	end


	def get__rels_dotrels(data) do
		{'_rels/.rels',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
		}
	end

	def get__rels_dir(data) do
		[get__rels_dotrels(data)]
	end

	def get_xl_rels_dir(data) do
		[{'xl/_rels/workbook.xml.rels',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>
"""
		}]
	end

	def get_xl_styles_xml(data) do
		{'xl/styles.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="GENERAL"/>
  </numFmts>
  <fonts count="1">
    <font>
      <sz val="10"/>
      <name val="Arial"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0"/>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle builtinId="0" name="Normal" xfId="0"/>
  </cellStyles>
</styleSheet>
"""	
		}
	end

	def get_xl_workbook_xml(data) do
		{'xl/workbook.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="Calc"/>
  <bookViews>
    <workbookView activeTab="0"/>
  </bookViews>
  <sheets>
    <sheet name="Sheet1" sheetId="1" state="visible" r:id="rId2"/>
  </sheets>
  <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>
</workbook>
"""
		}
	end

	def get_xl_sharedStrings_xml(data) do
		{'xl/sharedStrings.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">
  <si>
    <t>Hi</t>
  </si>
</sst>
"""
		}
	end

  @col_alphabet to_string(Enum.to_list(?A..?Z))
  
  @doc ~S"""
    returns the column letter(s) associated with a column index. Col idx starts at 1.
  """
  def encode_col(0) do "" end
  def encode_col num do
    znum = num - 1
    encode_col(div(znum, String.length(@col_alphabet))) <> String.at(@col_alphabet, rem(znum, String.length(@col_alphabet)))
  end

  @doc ~S"""
  Returns the Char/Number representation of a given row/column combination.
  Indizes start with 1.

  ## Examples
  
    iex> Elixlsx.to_excel_coords(1, 1)
    "A1"

    iex> Elixlsx.to_excel_coords(10, 27)
    "AA10"

  """
  def to_excel_coords(row, col) do
    encode_col(col) <> to_string(row)
  end

  defp xl_sheet_cols(row, rowidx) do
    Enum.zip(row, 1 .. length row) |>
    Enum.map(
      fn {col, colidx} ->
        {type, value} = cond do
          is_number(col) -> {"n", to_string(col)}
          true -> {"s", to_string(col)} # TODO this may throw.
        end
        List.foldr ["<c r=\"",
                     to_excel_coords(rowidx, colidx),
                     "\" s=\"0\" t=\"",
                     type,
                     "\">",
                     "<v>",
                     value,
                     "</v></c>"], "", &<>/2
        end) |>
    List.foldr "", &<>/2
  end

  defp xl_sheet_rows(data) do
    Enum.zip(data, 1 .. length data) |>
    Enum.map(fn {row, rowidx} -> 
        List.foldr(["<row r=\"",
                     to_string(rowidx),
                     "\">\n",
                     xl_sheet_cols(row, rowidx),
                     "</row>"], "", &<>/2) end) |>
    List.foldr("", &<>/2)
  end

	def get_xl_worksheets_dir(data) do
		[{'xl/worksheets/sheet1.xml',
			~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetPr filterMode="false">
    <pageSetUpPr fitToPage="false"/>
  </sheetPr>
  <dimension ref="A1"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="12.8"/>
  <sheetData>
  """ 
  <>
  xl_sheet_rows(data.rows)
  <>
  ~S"""
  </sheetData>
  <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
</worksheet>
			"""
			}]
	end

	def get_contentTypes_xml(data) do
		{'[Content_Types].xml',
			~S"""
<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>
"""
		}
	end

	def get_xl_dir(data) do
		[ get_xl_styles_xml(data),
			get_xl_sharedStrings_xml(data),
			get_xl_workbook_xml(data)] ++
		get_xl_rels_dir(data) ++
		get_xl_worksheets_dir(data)
	end

  def write_to(data, filename) do
    :zip.create(filename,
			get_docProps_dir(data) ++
      get__rels_dir(data) ++
      get_xl_dir(data) ++
      [ get_contentTypes_xml(data) ])
  end
end

