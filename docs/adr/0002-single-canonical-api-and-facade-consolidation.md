# ADR 0002: Single Canonical API Principle and Top-Level Facade Consolidation

* Status: Accepted
* Date: 2026-11-05
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb`](../../lib/xlsxrb.rb), [`Xlsxrb::Utils`](../../lib/xlsxrb/utils.rb), [`Xlsxrb::Elements::Cell`](../../lib/xlsxrb/elements/cell.rb), [`Xlsxrb::Elements::CellError`](../../lib/xlsxrb/elements/types.rb), [`Xlsxrb::Error`](../../lib/xlsxrb.rb)

---

## Context

Prior to version 0.2.0, xlsxrb accumulated multiple redundant helper namespaces and duplicate methods:
1. Coordinate conversions were scattered across `Xlsxrb::Utils` (`col_name_to_index`, `col_index_to_name`, `ref_to_row_col`, `row_col_to_ref`, `parse_range`, `range_to_ref`), `Elements::Cell` (`column_index`, `column_letter`, `parse_ref`), and top-level `Xlsxrb`.
2. Date/time serial calculations existed under both `Xlsxrb::Utils` and `Xlsxrb` with diverging naming (`datetime_to_serial`, `serial_to_datetime`, `date_to_serial`).
3. Excel cell error constants (`DIV0`, `NA`, `NAME`, `NULL`, `NUM`, `REF`, `VALUE`) were implemented as subclass constants under `Xlsxrb::Error < StandardError`, conflating exceptional runtime errors with normal spreadsheet cell values.
4. Top-level `Xlsxrb.format` shadowed Ruby core `Kernel#format` when called in singleton or included contexts.

This fragmentation increased maintenance overhead, complicated static type signatures (RBS/Steep), and created ambiguity for library consumers.

---

## Guiding Principles

1. Single Canonical API Principle:
   Every public capability in xlsxrb must have exactly one canonical method name. The codebase will not maintain permanent non-deprecated aliases for the same operation. Compatibility aliases intended to mirror other gems (such as roo or caxlsx) belong in an external adapter layer (`xlsxrb-adapters`), not in the core gem.
2. Explicit and Conflict-Free Names:
   Core library entry points must not shadow Ruby standard methods (`Kernel#format`).
3. Domain Separation:
   Cell error values represent domain data within a worksheet cell, not Ruby exceptions. They must be typed as domain value objects rather than subclasses of `StandardError`.

---

## Decision

### 1. Public Coordinate Conversion Facade
Consolidate all public coordinate helpers onto the top-level `Xlsxrb` module:
* `Xlsxrb.parse_ref(ref)`: Parses an A1 reference into 0-based `[row, column]`.
* `Xlsxrb.cell_to_ref(row, column)`: Converts 0-based `(row, column)` to an A1 string.
* `Xlsxrb.split_coordinate(ref)`: Splits an A1 reference into `[column_letter, 1_based_row]`.
* `Xlsxrb.column_index(col)`: Converts a column letter or symbol to a 0-based column index.
* `Xlsxrb.column_letter(index)`: Converts a 0-based column index to an uppercase column letter.
* `Xlsxrb.parse_range(range)`: Parses a range string into `[[top_row, left_col], [bottom_row, right_col]]`.
* `Xlsxrb.range_to_ref(...)`: Converts coordinates to a range string, accepting either 4 flat integers or nested coordinate pairs.
* `Xlsxrb.expand_formula(formula, ...)`: Shifts relative cell references within a formula expression.

Duplicate entry points on `Xlsxrb::Utils` (including `split_coordinate` and `expand_formula`) and `Elements::Cell` are deprecated for public use and marked `# @api private` for internal zero-allocation paths.

### 2. Date and Time Serial Normalization
* Canonical conversion methods are defined on `Xlsxrb`: `date_to_serial`, `time_to_serial`, `serial_to_date`, `serial_to_time`, `valid_excel_date?`.
* `datetime_to_serial` and `serial_to_datetime` are deprecated in favor of `time_to_serial` and `serial_to_time`.
* `Xlsxrb::Utils.valid_excel_date?` is deprecated in favor of `Xlsxrb.valid_excel_date?`.

### 3. Resolution of Kernel#format Shadowing
* Add `Xlsxrb.format_value(value, format_code = nil, date1904: false)` as the canonical number and date value formatting facade.
* Deprecate `Xlsxrb.format`. `NumberFormatter.format` remains unchanged.

### 4. Cell Error Value Objects
* Move cell error constants to frozen `Elements::CellError` value objects on `Elements::CellError`: `DIV0`, `NA`, `NAME`, `NULL`, `NUM`, `REF`, `VALUE`, and `GETTING_DATA`.
* Constants under `Xlsxrb::Error` are marked deprecated using `Object.deprecate_constant`.

### 5. Sheet Name Sanitization and Formula Reference Parameters
* `DslHelpers.sanitize_sheet_name(raw)` is the canonical helper for sanitizing sheet names against OpenXML length (31 chars) and invalid character constraints. `DslHelpers.sheet_name` is deprecated.
* `Xlsxrb.array_formula(expression, ref: ...)`: Positional parameter `pos_ref` is deprecated in favor of the explicit `ref:` keyword argument.

### 6. Builder Value Object Delegators
* `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` define direct delegator methods for the 5 conflict-free Value Object factories (`formula`, `array_formula`, `shared_formula`, `rich_text`, `cell_error`) that forward to `Xlsxrb.*` factory methods to construct canonical `Elements::*` Value Objects.
* `Xlsxrb.hyperlink` and `Xlsxrb.comment` remain module-level factories only on `Xlsxrb` to avoid semantic collisions with the worksheet-level DSL commands `hyperlink(ref, url, ...)` and `comment(ref, text, ...)`.
* Dead internal methods `build_raw_cell`, `build_raw_cell_from_value`, and `build_row_attrs` in `lib/xlsxrb.rb` have been removed.

---

## Consequences

### Positive
* Predictable Public Surface: Developers have a single documented location (`Xlsxrb.*`) for common utilities.
* Type Safety: Static RBS signatures map cleanly to distinct types without union confusion.
* Clean Exception Hierarchy: Rescuing `Xlsxrb::Error` catches actual runtime failures rather than capturing cell error values.

### Negative / Trade-offs
* Migration Notice: Callers referencing `Xlsxrb::Utils` or `Xlsxrb::Error::NA` receive runtime deprecation warnings in 0.1.x until updated for 0.2.0.
