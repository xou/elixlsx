defmodule Elixlsx.Util do
  @col_alphabet to_string(Enum.to_list(?A..?Z))

  @doc ~S"""
  returns the column letter(s) associated with a column index. Col idx starts at 1.

  ## Example

      iex> encode_col(1)
      "A"

      iex> encode_col(28)
      "AB"

  """
  @spec encode_col(non_neg_integer) :: String.t
  def encode_col(0), do: ""
  def encode_col(num) when num <= 26, do: <<num + 64>>
  def encode_col(num) do
    mod = div(num, 26)
    rem = rem(num, 26)

    if rem == 0 do
      encode_col(mod - 1) <> encode_col(26)
    else
      encode_col(mod) <> encode_col(rem)
    end
  end

  @doc ~S"""
  returns the column index associated with a given letter.

  ## Example

      iex> decode_col("AB")
      28

      iex> decode_col("A")
      1

  """
  @spec decode_col(list(char()) | String.t) :: non_neg_integer
  def decode_col s do
    cond do
      is_list s -> decode_col(to_string s)
      String.valid? s -> decode_col_ s
      true -> raise %ArgumentError{message: "decode_col expects string or charlist, got "
                                   <> inspect s}
    end
  end


  @spec decode_col_(String.t) :: non_neg_integer
  defp decode_col_("") do 0 end
  defp decode_col_(s) do
    alphabet_list = String.to_charlist @col_alphabet

    if !String.match? s, ~r/^[A-Z]*$/ do
      raise %ArgumentError{message: "Invalid column string: " <> inspect s}
    end

    # translate list of strings to the base-26 value they represent
    Enum.map(String.to_charlist(s), (fn x -> :string.chr(alphabet_list, x) end)) |>
    # multiply and aggregate them
    List.foldl(0, (fn (x, acc) -> x + 26 * acc end))
  end


  @doc ~S"""
  Returns the Char/Number representation of a given row/column combination.
  Indizes start with 1.

  ## Examples

      iex> to_excel_coords(1, 1)
      "A1"

      iex> to_excel_coords(10, 27)
      "AA10"

  """
  @spec to_excel_coords(number, number) :: String.t
  def to_excel_coords(row, col) do
    encode_col(col) <> to_string(row)
  end

  @spec from_excel_coords(String.t) :: {pos_integer, pos_integer}
  @doc ~S"""
  returns a tuple {row, col} corresponding to the input.
  row and col are 1-indexed, use from_excel_coords0 for zero-indexing.

  Example:

      iex> from_excel_coords("C2")
      {2, 3}

      iex> from_excel_coords0("C2")
      {1, 2}

  """
  def from_excel_coords(input) do
    case Regex.run(~r/^([A-Z]+)([0-9]+)$/, input, capture: :all_but_first) do
      nil -> raise %ArgumentError{
                    message: "Invalid excel coordinates: " <>
                            (inspect input)}
      [colS, rowS] ->
        {row, _} = Integer.parse rowS
        {row, decode_col(colS)}
    end
  end

  @spec from_excel_coords0(String.t) :: {non_neg_integer, non_neg_integer}
  @doc ~S"See from_excel_coords/1"
  def from_excel_coords0(input) do
    {row, col} = from_excel_coords(input)
    {row - 1, col - 1}
  end


  @doc ~S"""
  Returns the ISO String representation (in UTC) for a erlang datetime() or datetime1970()
  object.

  ## Example

      iex> iso_from_datetime {{2000, 12, 30}, {23, 59, 59}}
      "2000-12-30T23:59:59Z"

  """
  @type datetime_t :: :calendar.datetime()
  @spec iso_from_datetime(datetime_t) :: String.t
  def iso_from_datetime calendar do
    {{y, m, d}, {hours, minutes, seconds}} = calendar
    to_string(:io_lib.format('~4.10.0b-~2.10.0b-~2.10.0bT~2.10.0b:~2.10.0b:~2.10.0bZ',
                             [y, m, d, hours, minutes, seconds]))
  end


  @doc ~S"""
  returns

  - the current current timestamp if input is nil,
  - the UNIX-Timestamp interpretation when given an integer,

  both in ISO-Repr.

  If input is a String, the string is returned.

      iex> iso_timestamp 0
      "1970-01-01T00:00:00Z"

      iex> iso_timestamp 1447885907
      "2015-11-18T22:31:47Z"

  It doesn't validate string inputs though:

      iex> iso_timestamp "goat"
      "goat"

  """
  @spec iso_timestamp(String.t | integer | nil) :: String.t
  def iso_timestamp input \\ nil do
    cond do
      input == nil ->
        iso_from_datetime(:calendar.universal_time)
      is_integer(input) ->
        iso_from_datetime(:calendar.now_to_universal_time({div(input, 1000000), rem(input, 1000000), 0}))
      # TODO this case should parse the string i guess
      # TODO also prominently absent: [char].
      String.valid? input ->
        input
      true -> raise "Invalid input to iso_timestamp." <> (inspect input)
    end
  end


  @excel_epoch {{1899, 12, 31}, {0, 0, 0}}
  @secs_per_day 86400

  @doc ~S"""
  Convert an erlang :calendar object to an excel timestamp.
  """
  @spec to_excel_datetime(datetime_t) :: {:excelts, number}
  def to_excel_datetime({{yy, mm, dd}, {h, m, s}}) do
    in_seconds = :calendar.datetime_to_gregorian_seconds {{yy, mm, dd}, {h, m, s}}
    excel_epoch = :calendar.datetime_to_gregorian_seconds @excel_epoch

    t_diff = (in_seconds - excel_epoch) / @secs_per_day

    # Apply the "Lotus 123" bug - 1900 is considered a leap year.
    t_diff = if t_diff > 59 do
      t_diff + 1
    else
      t_diff
    end

    {:excelts, t_diff}
  end


  @doc ~S"""
  Convert a unix timestamp to excel time.
  """
  @spec to_excel_datetime(number) :: {:excelts, number}
  def to_excel_datetime(input) when is_number(input) do
    to_excel_datetime(
      :calendar.now_to_universal_time({div(input, 1000000), rem(input, 1000000), 0}))
  end

  @doc ~S"""
  Timestampts that are already in excel format are passed through
  unmodified.
  """
  @spec to_excel_datetime({:excelts, number}) :: {:excelts, number}
  def to_excel_datetime({:excelts, value}) do
    {:excelts, value}
  end

  @doc ~S"""
  Formula's value calculate on opening excel programm. We don't need to format this here.
  """
  @spec to_excel_datetime({:formula, String.t}) :: {:formula, String.t}
  def to_excel_datetime({:formula, value}) do
    {:formula, value}
  end

  @doc ~S"""
  replace_all(input, [{search, replace}])

  ## Example

      iex> replace_all("Hello World", [{"e", "E"}, {"o", "oO"}])
      "HElloO WoOrld"

  """
  @spec replace_all(String.t, [{String.t, String.t}]) :: String.t

  def replace_all(input, [{s,r}|srx]) do
    String.replace(input, s, r) |> replace_all(srx)
  end

  def replace_all(input, []) do
    input
  end


  @version Mix.Project.config[:version]
  @doc ~S"""
  Returns the application version suitable for the <ApplicationVersion> tag.
  """
  def app_version_string do
    String.replace(@version, ~r/(\d+)\.(\d+)\.(\d+)/, "\\1.\\2\\3")
  end
end

