defmodule Elixlsx.XML_Templates do
  alias Elixlsx.Util, as: U

  @docprops_app ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <TotalTime>0</TotalTime>
  <Application>Elixlsx</Application>
  <AppVersion>0.0.1</AppVersion>
</Properties>
"""
  def docprops_app, do: @docprops_app


  @docprops_core ~S"""
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dcterms:created xsi:type="dcterms:W3CDTF">__TIMESTAMP__</dcterms:created>
  <dc:language>__LANGUAGE__</dc:language>
  <dcterms:modified xsi:type="dcterms:W3CDTF">__TIMESTAMP__</dcterms:modified>
  <cp:revision>__REVISION__</cp:revision>
</cp:coreProperties>
"""

  def docprops_core(timestamp \\ nil, language \\ "en-US", revision \\ 1) do
    timestamp_ = U.iso_timestamp(timestamp)
    @docprops_core |>
    String.replace("__TIMESTAMP__", timestamp_) |>
    String.replace("__LANGUAGE__", language) |>
    String.replace("__REVISION__", to_string(revision))
  end 
end
