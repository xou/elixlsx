# Changelog

## 0.6.0 [WIP]

! Minimum elixir version raised to ~~1.7 (#116)~~ 1.12 (#140)
- Add cell validations (#109)
- Fix some typos (#120)
- Add a new YYYY-MM date format (#123)
- Fix warnings for range syntax (#139)
- Update README.md, call out requirement to escape '<' and '>' characters (#138)
- Fix: Actually throw ArgumentError when cell content type is invalid (#140)

## 0.5.1

- Revert autofilter support (breaks build)
- Migrate travis-ci.com, fixing the CI system (which apparently failed open for months) 

## 0.5.0

- Autofilter support (#105)
- Namespace XML module (#111, #107)
- Documentation improvements (#103, #106)
- Preserve leading whitespace (#95)
- `mix format` the entire codebase

## 0.4.2

- Dialyzer fixes
- Speed improvements
- XML character validation fixes

## 0.4.1

- Fixes for Dialyzer, documentation.
- Added support for boolean values in cells

## 0.4.0

### Changed

- Increased minimum elixir version required to 1.3
