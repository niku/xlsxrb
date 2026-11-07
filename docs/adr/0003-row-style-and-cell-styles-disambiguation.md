# ADR 0003: Row Style and Cell Styles Disambiguation

* Status: Accepted
* Date: 2026-11-05
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Elements::Row`](../../lib/xlsxrb/elements/row.rb), [`Xlsxrb::Elements::Cell`](../../lib/xlsxrb/elements/cell.rb)

---

## Context

In xlsxrb 0.1, row generation methods (`WorksheetBuilder#row`, `StreamWriter#row`, `Row.new`) accepted two keywords with only a single-character difference:
1. `style:` specified the OpenXML row-level default formatting applied to the `<row s="..." customFormat="1">` element.
2. `styles:` specified formatting applied to individual cell `<c s="...">` elements, either as a single style broadcast to every cell in the row or as an Array of styles aligned by column index.

This 1-character difference produced subtle defects:
* A developer intending to style row cells by writing `sheet.row(data, style: :header)` inadvertently formatted only the `<row>` container. In certain spreadsheet engines (including LibreOffice Calc and web viewers), row formatting does not cascade to explicitly populated `<c>` elements, resulting in missing cell borders, backgrounds, and fonts.
* Passing an Array to `style:` triggered confusing runtime errors or unexpected behavior.
* In `Elements::Cell`, the `style:` parameter was overloaded to accept either a symbolic name (`:bold`), a style Hash (`{ font: { bold: true } }`), or a 0-based stylesheet integer index (`0`, `1`, `2`).

---

## Guiding Principles

1. Unambiguous Parameter Identifiers:
   Keywords must clearly reflect whether their target is the row container itself or the individual cells within the row.
2. Type Separation:
   Symbolic or declarative style specifications must not be conflated with low-level stylesheet index integers.

---

## Decision

### 1. Rename to cell_styles: and row_style:
On row construction methods across `WorksheetBuilder`, `StreamWriter`, and `Elements::Row`:
* Rename `styles:` to `cell_styles:`. Accepts an Array of styles (`[s1, s2, ...]`) or a single style Symbol/Hash broadcast across all cells.
* Rename `style:` to `row_style:`. Accepts a single style Symbol/Hash applied to the `<row>` element.
* Retain deprecated `styles:` and `style:` keywords with runtime warnings directing callers to `cell_styles:` and `row_style:`.

### 2. Disambiguation in Elements::Cell and Promotion in Elements::Row
In `Elements::Cell`:
* `style:` is reserved for declarative styles (Symbol, Hash, or `Elements::Style`).
* `style_index:` (or `xf_index:`) is reserved for the 0-based Integer index into the stylesheet `<cellXfs>`.

In `Elements::Row`:
* `:style_index` is promoted to a declared member of `Elements::Row` (`Data.define`), migrating row stylesheet indices out of `unmapped_data[:style_index]`. `WorksheetBuilder#row` and `build_row_from_raw` initialize rows with `style_index:`.

---

## Consequences

### Positive
* Clarity: Callers immediately recognize whether formatting applies to cells (`cell_styles:`) or the row container (`row_style:`).
* Bug Prevention: Eliminates accidental row container formatting when cell formatting was intended.
* Type Safety: Clear separation between declarative styles and low-level stylesheet integers simplifies static RBS signatures.

### Negative / Trade-offs
* Migration Effort: Callers using `sheet.row(vals, styles: ...)` must update to `sheet.row(vals, cell_styles: ...)`. Deprecation warnings guide this transition before 0.2.0 removal.
