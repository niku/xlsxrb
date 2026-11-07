# ADR 0005: DSL Command Void Return, Property Accessors, and Predicate Conventions

* Status: Accepted
* Date: 2026-11-05
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::WorkbookBuilder`](../../lib/xlsxrb/workbook_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Elements`](../../lib/xlsxrb/elements.rb)

---

## Context

Prior to version 0.2.0, xlsxrb builder methods had inconsistent return values:
1. Some DSL methods in `WorksheetBuilder` returned intermediate data structures (e.g. `sheet.row(...)` returned the appended `Row` or Hash, `sheet.merge(...)` returned a merged array), while corresponding methods in `StreamWriter` returned `nil` because rows were immediately streamed to disk. This divergence prevented transparent interchangeability between in-memory and streaming builders.
2. Property configuration methods were ambiguous: calling `sheet.page_setup(...)` with arguments acted as a mutator, but calling without arguments was either undefined or collided with setting properties to `nil`.
3. Boolean state inspection was inconsistent across objects: some used bare nouns (`row.hidden`, `cell.link`), while others used Ruby-idiomatic predicates (`row.hidden?`, `cell.hyperlink?`).

---

## Guiding Principles

1. Streaming/In-Memory Builder Interchangeability:
   Any block of code written against `WorksheetBuilder` must be executable against `StreamWriter` and `WorksheetProxy` without behavior divergence.
2. Command-Query Separation in DSLs:
   Imperative builder actions are commands that return `void` (`nil`). The only exceptions are explicit chaining operators (`#<<`) and dedicated fluent configurators (`StyleBuilder`).
3. Standard Ruby Predicate Convention:
   Methods querying boolean state on read-only value objects (`Elements::*`), streaming readers (`StreamSheet`), and row wrappers (`StreamRow`) must end with `?`.

---

## Decision

### 1. Void Return for DSL Commands
All builder mutation commands return `nil`:
* `WorksheetBuilder`, `WorkbookBuilder`, `StreamWriter`, and `WorksheetProxy`: `row`, `merge`, `hyperlink`, `comment`, `image`, `table`, `chart`, `shape`, `data_validation`, `conditional_format`, `freeze_pane`, `split_pane`, `row_break`, `column_break`, `header_footer`, `page_setup`, `page_margins`, `print_options`, `sheet_protection`, `workbook_protection`.
* `StyleBuilder` also follows the `void` command + 0-arg getter / `=` setter pattern.
* Contract tests (`test/contract/api_symmetry_test.rb`) enforce identical arity and `nil` return types across both builder implementations.
* Fluent chaining `-> self` is reserved exclusively for the `#<<` row appending operator (and `StyleBuilder#<<`).

### 2. Setter and Getter Symmetry
For configuration properties on builders:
* Attribute writer `sheet.property = val` is the canonical assignment.
* Single-value properties (`print_area`, `auto_filter`, `code_name`, `tab_color`, `fit_to_page`, `state`, `background_image`, `phonetic_properties`, `sheet_calc_properties`) use 0-argument Getter + `=` Setter conventions.
* Structured multi-option properties (`page_setup`, `page_margins`, `header_footer`, `sheet_protection`, `workbook_protection`, `freeze_pane`, `split_pane`, `selection`, `sort_state`, `print_titles`, `outline_properties`, `sheet_view`, `sheet_format`, `sheet_properties`, `print_options`):
  * When called with 0 arguments, act as getters returning their corresponding immutable `Elements::*` Value Object (or `nil` / Hash slice for outline properties).
  * When called with positional or keyword arguments or via `=` setter, act as mutators and return `void` (`nil`) (or the assigned value for `=`). Setters (`sheet_protection=`, `workbook_protection=`, `print_titles=`, `page_setup=`, etc.) accept either a `Hash` or the corresponding Value Object (or boolean/nil where applicable).
* `columns`: 0-argument reader on `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, and `TableBuilder` returning an array of `Elements::Column` (or `Elements::TableColumn`), while `column(...)` commands return `void` (`nil`).
* `TableBuilder`: Hash-like access via `#[]` and `#[]=` is deprecated in favor of explicit getters and setters (`#style`, `#style=`, `#show_row_stripes`, `#show_row_stripes=`, `#show_column_stripes`, `#show_column_stripes=`, etc.).
* `StyleBuilder`: Follows the `void` command + 0-arg getter / `=` setter pattern, providing 0-argument getters (`font_name`, `font_color`, `fill_color`, `align_horizontal`, `align_vertical`, `text_rotation`, `indent`, `vert_align`) and boolean query predicates (`bold?`, `italic?`, `strike?`, `underline?`, `wrap_text?`, `shrink_to_fit?`, `locked?`, `hidden?`).

### 3. Boolean Predicate Standardization
* Readers and value objects standardize on `?` predicates: `row.hidden?`, `row.custom_height?`, `row.collapsed?`, `cell.comment?`, `cell.hyperlink?`, `cell.formula?`, `workbook.date1904?`.
* Value objects in `Elements::*` (`MetadataRecord`) expose boolean `?` predicates returning `bool` (`#: () -> bool`):
  * `Font`: `bold?`, `italic?`, `strike?`, `outline?`, `shadow?`, `condense?`, `extend?`
  * `Alignment`: `wrap_text?`, `shrink_to_fit?`, `justify_last_line?`
  * `Protection`: `locked?`, `hidden?`
  * `CellXf`: `quote_prefix?`, `pivot_button?`, `apply_number_format?`, `apply_font?`, `apply_fill?`, `apply_border?`, `apply_alignment?`, `apply_protection?`
  * `PageSetup`: `black_and_white?`, `draft?`
  * `HeaderFooter`: `different_first?`, `different_odd_even?`, `scale_with_doc?`, `align_with_margins?`
  * `PrintOptions`: `grid_lines?`, `grid_lines_set?`, `headings?`, `horizontal_centered?`, `vertical_centered?`
  * `SheetProperties`: `fit_to_page?`, `apply_styles?`, `summary_below?`, `summary_right?`, `show_outline_symbols?`
  * `SheetFormat`: `custom_height?`, `zero_height?`, `thick_top?`, `thick_bottom?`
  * `SheetView`: `window_protection?`, `show_formulas?`, `show_grid_lines?`, `show_row_col_headers?`, `show_zeros?`, `right_to_left?`, `tab_selected?`, `show_ruler?`, `show_outline_symbols?`, `default_grid_color?`, `show_white_space?`
  * `SheetCalcProperties`: `full_calc_on_load?`
  * `PageBreak`: `manual?`
  * `DefinedName`: `hidden?`
  * `WorkbookProperties`: `date1904?`, `filter_privacy?`
  * `WorkbookProtection`: `lock_structure?`, `lock_windows?`, `lock_revision?`
  * `SheetProtection`: `sheet?`, `objects?`, `scenarios?`, `format_cells?`, `format_columns?`, `format_rows?`, `insert_columns?`, `insert_rows?`, `insert_hyperlinks?`, `delete_columns?`, `delete_rows?`, `select_locked_cells?`, `sort?`, `auto_filter?`, `pivot_tables?`, `select_unlocked_cells?`
  * `FilterColumn`: `hidden?`, `show_button?`
  * `Comment`: `visible?`
  * `ThreadedComment`: `resolved?`, `done?`
  * `DataValidation`: `allow_blank?`, `show_input_message?`, `show_error_message?`, `show_drop_down?`
  * `Table`: `totals_row_shown?`, `header_row_shown?`
  * `SparklineGroup`: `markers?`, `high?`, `low?`, `first?`, `last?`, `negative?`
  * `PivotTable`: `column_grand_totals?`, `row_grand_totals?`, `show_row_headers?`, `show_column_headers?`, `show_row_stripes?`, `show_column_stripes?`
  * `IgnoredError`: `eval_error?`, `two_digit_text_year?`, `number_stored_as_text?`, `formula?`, `formula_range?`, `unlocked_formula?`, `empty_cell_reference?`, `list_data_validation?`, `calculated_column?`
* Structural element boolean query predicates:
  * `Row` and `StreamRow`: `custom_height?`
  * `Column`: `custom_width?`, `best_fit?`, `collapsed?`
* Bare property readers (`row.hidden`, `cell.link`) are deprecated with warnings.

---

## Consequences

### Positive
* Perfect Symmetry: Code authored for in-memory workbooks can be swapped to streaming writers without return-value dependencies breaking the pipeline.
* Idiomatic Ruby: Consistent `?` predicates on elements and inspection interfaces.
* Static Typability: Clean RBS annotations (`#: (...) -> void`) without complex union return types.

### Negative / Trade-offs
* Elimination of Accidental Chaining: Callers relying on `sheet.row(...)` returning the newly created `Row` must access rows via `sheet.rows.last` or `sheet[row_idx]` in in-memory builders.
