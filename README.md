# Elixlsx

[![Hex Version](http://img.shields.io/hexpm/v/elixlsx.svg?style=flat)](https://hex.pm/packages/elixlsx)

A writer for XLSX files.

Features:

- Multiple (named) sheets with custom column widths & column heights.
- (Unicode-)strings, Numbers, Dates
- Font formatting (size, bold, italic, underline, strike)
- Horizontal alignment and text wrapping
- Font and cell background color, borders
- Merged cells

## Installation

Via hex, in mix.exs:

```Elixir
defp deps do
  [{:elixlsx, "~> 0.2.0"}]
end
```

Via github:

```Elixir
defp deps do
  [{:elixlsx, git: "https://github.com/xou/elixlsx.git"}]
end
```

## Usage

1-Line tutorial:

```Elixir
(alias Elixlsx.Workbook, alias Elixlsx.Sheet)
iex(1)> Workbook.append_sheet(%Workbook{}, Sheet.with_name("Sheet 1") |> Sheet.set_cell("A1", "Hello", bold: true)) |> Elixlsx.write_to("hello.xlsx")
```

See [example.exs](example.exs) for examples how to use the various features.

## Number and date formatting reference

A quick introduction how number formattings look like can be found
[here](https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match)

