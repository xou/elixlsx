defmodule Elixlsx.Style.NumFmt do
  alias Elixlsx.XML
  alias __MODULE__

  defstruct format: nil

  @type t :: %NumFmt{
          format: String.t()
        }

  def from_props(props) do
    cond do
      props[:yyyymmdd] ->
        date_yyyy_mm_dd()

      props[:yyyymm] ->
        date_yyyy_mm()

      props[:datetime] ->
        date_datetime()

      props[:num_format] ->
        if XML.valid?(props[:num_format]) do
          %NumFmt{format: props[:num_format]}
        else
          raise %ArgumentError{
            message:
              ":date_format parameter needs to be a" <>
                " valid string, but is " <> inspect(props[:num_format])
          }
        end

      true ->
        nil
    end
  end

  @doc ~S"""
  Returns a semi-ISO date format (YYYY-MM-DD).
  """
  def date_yyyy_mm_dd do
    %NumFmt{format: "yyyy-mm-dd"}
  end

  @doc ~S"""
  Returns a semi-ISP date format only month and year (YYYY-MM)
  """
  def date_yyyy_mm do
    %NumFmt{format: "yyyy-mm"}
  end

  @doc ~S"""
  Returns a standard semi-iso datetime format.
  """
  def date_datetime do
    %NumFmt{format: "yyyy-mm-dd h:mm:ss"}
  end

  def is_date?(numfmt) do
    # TODO while this is probably reliable enough for 99% of cases...
    String.contains?(numfmt.format, "yy")
  end

  @doc ~S"""
  Returns a style XML entry.

  In contrast to most (all?) other entries in styles.xml,
  numFmts actually carry their index in each element; probably
  to avoid confusion with the built-in styles. Therefore,
  get_stylexml_entry requires a second parameter here.
  """
  def get_stylexml_entry(numfmt, idx) do
    fmt = Elixlsx.XMLTemplates.xml_escape(numfmt.format)

    """
    <numFmt numFmtId="#{idx}" formatCode="#{fmt}" />
    """
  end
end
