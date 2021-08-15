# Elixlsx

[![Build Status](https://travis-ci.com/xou/elixlsx.svg?branch=master)](https://travis-ci.org/xou/elixlsx)
[![Module Version](https://img.shields.io/hexpm/v/elixlsx.svg)](https://hex.pm/packages/elixlsx)
[![Hex Docs](https://img.shields.io/badge/hex-docs-lightgreen.svg)](https://hexdocs.pm/elixlsx/)
[![Total Download](https://img.shields.io/hexpm/dt/elixlsx.svg)](https://hex.pm/packages/elixlsx)
[![License](https://img.shields.io/hexpm/l/elixlsx.svg)](https://github.com/xou/elixlsx/blob/master/LICENSE)
[![Last Updated](https://img.shields.io/github/last-commit/xou/elixlsx.svg)](https://github.com/xou/elixlsx/commits/master)

Elixlsx is a writer for the MS Excel OpenXML format (`.xlsx`).

Features:

- Multiple (named) sheets with custom column widths & column heights.
- (Unicode-)strings, Numbers, Dates
- Font formatting (size, bold, italic, underline, strike)
- Horizontal alignment and text wrapping
- Font and cell background color, borders
- Merged cells

## Installation


### Elixir 1.3 and above

Via Hex, in `mix.exs`:

```elixir
defp deps do
  [{:elixlsx, "~> 0.5.1"}]
end
```

Via GitHub:

```elixir
defp deps do
  [{:elixlsx, github: "xou/elixlsx"}]
end
```

### Elixir pre 1.3

```elixir
defp deps do
  [{:elixlsx, "~> 0.3.1"}]
end
```

Via github:

```elixir
defp deps do
  [{:elixlsx, github: "xou/elixlsx", tag: "v0.3.1"}]
end
```

## Usage

1-Line tutorial:

```elixir
(alias Elixlsx.Workbook, alias Elixlsx.Sheet)
iex(1)> Workbook.append_sheet(%Workbook{}, Sheet.with_name("Sheet 1") |> Sheet.set_cell("A1", "Hello", bold: true)) |> Elixlsx.write_to("hello.xlsx")
```

See [example.exs](https://github.com/xou/elixlsx/blob/master/example.exs) for examples how to use the various features.

## Number and date formatting reference

A quick introduction how number formattings look like can be found
[here](https://social.msdn.microsoft.com/Forums/office/en-US/e27aaf16-b900-4654-8210-83c5774a179c/xlsx-numfmtid-predefined-id-14-doesnt-match).


## License

Copyright (c) 2015 Nikolai Weh

This library is MIT licensed. See the [LICENSE](https://github.com/xou/elixlsx/blob/master/LICENSE) for details.
