# Elixlsx

A writer for XLSX files.

Supports:

- (Unicode-)strings, Numbers
- Font formatting (size, bold, italic, underline, strike)
- Multiple (named) sheets.

This library is currently more in a proof-of-concept state;
it is also my first Elixir project, feedback is very welcome.

## Usage

1-Line tutorial:

```Elixir
iex(1)> Workbook.append_sheet(%Workbook{}, Sheet.with_name("Sheet 1") |> Sheet.set_cell("A1", "Hello", bold: true)) |> Elixlsx.write_to("hello.xlsx")
```

See Font.from_props in elixlsx/style/font.ex for a full list
of currently supported formatting options.

## Number formatting reference

A quick introduction how nubmer formattings look like can be found [here](https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match)

