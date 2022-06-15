defmodule Elixlsx.XML do
  @xml_chars_block_1 [9, 10, 13]
  @xml_chars_block_2 32..55_295
  @xml_chars_block_3 57_344..65_533
  @xml_chars_block_4 65_536..1_114_111

  # From the xml spec 1.0: https://www.w3.org/TR/REC-xml/#charsets
  # Character Range
  #   any Unicode character, excluding the surrogate blocks, FFFE, and FFFF.
  #   Char ::= #x9 | #xA | #xD | [#x20-#xD7FF] | [#xE000-#xFFFD] | [#x10000-#x10FFFF]
  def valid?(<<h::utf8, t::binary>>)
      when h in @xml_chars_block_1 or h in @xml_chars_block_2 or h in @xml_chars_block_3 or
             h in @xml_chars_block_4 do
    valid?(t)
  end

  def valid?(<<>>), do: true
  def valid?(_), do: false
end
