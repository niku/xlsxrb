# ADR 0004: Uniform Zero-Based Indexing, Builder/Reader Symmetry, and Column Spelling

* Status: Accepted
* Date: 2026-11-05
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::DslHelpers`](../../lib/xlsxrb/dsl_helpers.rb), [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::WorkbookBuilder`](../../lib/xlsxrb/workbook_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Elements::Worksheet`](../../lib/xlsxrb/elements/worksheet.rb), [`Xlsxrb::StreamSheet`](../../lib/xlsxrb/stream_sheet.rb)

---

## Context

Prior to version 0.2.0, xlsxrb exhibited three architectural inconsistencies across its API surface:
1. Coordinate Indexing Inconsistency:
   Numeric indices across most of xlsxrb were 0-based (`sheet[0, 0]`, `freeze_pane(row: 1, column: 0)`, `row_break(0)`, `image(from_row: 0)`). However, `rows:` and `except_rows:` keyword arguments in `data_validation(rows: 2..100)`, `conditional_format(rows: 2..100)`, `defined_name(rows: 1..10)`, and `print_titles(rows: "1:2")` were 1-based (reflecting visual Excel row numbers). Developers constantly had to recall whether a given method accepted 0-based or 1-based integers.
2. Grammar and Grammatical Number Asymmetry:
   Method names mixed singular and plural forms arbitrarily. For instance, builders had plural `column_breaks` alongside singular `row_break`, while `conditional_formatting` mirrored the raw XML tag rather than matching `data_validations` and `conditional_formats`.
3. Column Abbreviation Fragmentation:
   The codebase mixed abbreviated `col`, `cols`, and full `column`, `columns` inconsistently across builders and readers (`col_breaks` vs `column_breaks`, `from_col` vs `from_column`, `default_col_width` vs `default_column_width`).

---

## Guiding Principles

1. Uniform Coordinate Rule:
   All numeric row and column parameters in Ruby methods are strictly 0-based (row 0 is Excel row 1, column 0 is Excel column A). String A1 references (`"A1:C10"`, `"1:2"`, `"A:B"`) represent standard spreadsheet notation and remain 1-based.
2. Singular Builder, Plural Reader Rule:
   DSL methods that construct a single element use the singular noun (e.g. `data_validation`, `conditional_format`, `ignored_error`, `row_break`, `column_break`). Reader methods that query collections use the plural noun (e.g. `data_validations`, `conditional_formats`, `ignored_errors`, `row_breaks`, `column_breaks`).
3. Explicit Naming:
   Standardize on full English spelling `column` across public methods and options, treating `col` as a deprecated alias.

---

## Decision

### 1. 0-Based Numeric Indexing for rows:, except_rows:, and Coordinate Access Boundaries
In `DslHelpers`:
* `DslHelpers.area_ref(col_run, row_run)` treats `r1, r2` as 0-based row indices and emits `r1 + 1` and `r2 + 1` in the resulting A1 range.
* `normalize_data_validation` and `normalize_conditional_format` default full-column ranges to `0...Xlsxrb::MAX_ROW`.
* `normalize_defined_name_value` treats integer/Range `rows` as 0-based (`0` -> Excel row `"1"`).
* `print_titles` accepts 0-based integer or Range `rows:` (`0..1` -> `"1:2"`) and 0-based integer, Symbol, or Range `columns:` (`0..1` or `:A..:B` -> `"A:B"`), alongside string ranges (`"1:2"`, `"A:B"`).

In `CoordinateAccess` and `StreamSheet`:
* `dimension_bounds` returns 0-based indices `{ min_row: r1, max_row: r2, min_column: c1, max_column: c2 }` instead of 1-based bounds.
* `row_range` returns 0-based `Range[Integer]` bounds (e.g. `"A1:C5"` -> `0..4`) directly composable with `sheet.row_at(r)` and `sheet.each_row(rows: sheet.row_range)`.
* `column_range` returns 0-based `Range[Integer]` bounds (e.g. `"A1:C5"` -> `0..2`) directly composable with `row.cell_at(c)`.

### 2. Singular Builder / Plural Reader Symmetry
* Standardize on singular builders: `data_validation`, `conditional_format`, `ignored_error`, `row_break`, `column_break`, `sheet_protection`, `workbook_protection`, `sparkline_group`.
* Plural reader properties: `sheet.data_validations`, `sheet.conditional_formats`, `sheet.ignored_errors`, `sheet.row_breaks`, `sheet.column_breaks`.
* Deprecate verb-style builders (`validate_data`, `protect_sheet`, `protect_workbook`, `ignore_error`) and XML-literal builders (`conditional_formatting`).

### 3. Full column Spelling
* Adopt `column`, `columns`, `from_column`, `to_column`, `default_column_width`, `outline_level_column`, `column_breaks`.
* Deprecate abbreviations `col`, `cols`, `from_col`, `to_col`, `default_col_width`, `outline_level_col`, `col_breaks`.
* `Elements::FilterColumn`: Declare `:column_id` as the canonical member (`Data.define`). Constructor keyword `col_id:` and reader `#col_id` are deprecated in favor of `column_id:` and `#column_id`.

### 4. Chart Axis Structures
* `Elements::Chart`: Expose category and value axes via structured `category_axis` and `value_axis` Hashes (`title:`, `orientation:`, etc.).
* Bare attribute readers `cat_axis_title` and `val_axis_title` are deprecated in favor of `category_axis[:title]` and `value_axis[:title]`.

### 5. Worksheet Calculation and Image Properties
* `Elements::Worksheet`: Use `sheet_calc_properties:` as the canonical parameter in `initialize` and `#with`. Keywords `sheet_calc_pr:`, `calc_pr:`, `phonetic_pr:`, `background_picture:`, and `picture:` are deprecated in favor of `sheet_calc_properties:`, `phonetic_properties:`, and `background_image:`.

### 6. Active Sheet Properties
* `Elements::WorkbookProperties`: Rename `:active_tab` member to `:active_sheet`.
* `Elements::Workbook`: Expose `#active_sheet_index` and `#active_sheet`. Getter `#active_tab` is deprecated in favor of `#active_sheet_index` and `#active_sheet`.

---

## Consequences

### Positive
* Predictability: Zero cognitive overhead when deciding whether to pass 0-based or 1-based integers to any numeric method in xlsxrb.
* Symmetry: Systematic pattern between singular creation (`s.conditional_format(...)`) and plural inspection (`s.conditional_formats`).
* Readability: Consistent `column` keyword naming across all builders.

### Negative / Trade-offs
* Migration: Callers passing 1-based row numbers to `rows: 2..10` in `data_validation` or `conditional_format` must subtract 1 (`rows: 1..9`). Deprecation warnings and migration documentation assist user upgrades.
