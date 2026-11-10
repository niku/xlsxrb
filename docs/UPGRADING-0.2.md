# Upgrading from 0.1.x to 0.2.0

0.2.0 removes the APIs that 0.1.x deprecated and changes a few defaults. This guide covers the changes most 0.1.x code runs into. The complete list is in the `0.2.0` section of [CHANGELOG.md](../CHANGELOG.md).

A practical way to upgrade: run your code on 0.1.18 first and fix every deprecation warning, then switch to 0.2.0 and use the tables below for whatever still fails. Read the next section first, because some 0.1.18 warnings pointed to methods that 0.2.0 then removed.

## 0.1.18 advice that 0.2.0 reversed

| 0.1.18 suggested | Use in 0.2.0 |
| :--- | :--- |
| `fit_to_page` instead of `fit_to_page?` on builders | `fit_to_page?` to read, `fit_to_page = true` to set |
| `StyleBuilder#bold(true)` (and `italic`, `wrap_text`, ...) | `style.bold = true`, or `style(:name, bold: true)`; read with `bold?` |
| `TableBuilder#add_column` | `table.column(name, **opts)` |
| `Xlsxrb.modify` → `Xlsxrb.load`, mutate, `Xlsxrb.write` | There is no in-place cell update (`update_cell` and `update_sheet` are gone). Stream with `Xlsxrb.read` into `Xlsxrb.write` (see the README section "Constant-Memory Streaming Transformation"), or build new `Elements::Worksheet`/`Workbook` objects with `#with` |

## Reading

| 0.1.x | 0.2.0 |
| :--- | :--- |
| `Xlsxrb.read(path).load`, `workbook.load` | `Xlsxrb.load(path)` |
| `Xlsxrb.read(path)` without a block, for random access | `Xlsxrb.load(path)`; blockless `Xlsxrb.read` now returns an `Enumerator` of `StreamSheet` |
| `sheet.cell("A1")` | `sheet["A1"]` |
| `row.to_a` for values | `row.values` (`to_a` now returns the `Cell` objects) |
| `stream_row.to_h` | `row.values` or `row.cells` |
| `cell.content` | `cell.value` |
| `sheet.first_row`, `sheet.last_row` | `sheet.row_range` or `sheet.dimension_bounds` (both come from the file's `<dimension>` and are `nil` without it) |
| `sheet.first_column`, `sheet.last_column` | `sheet.column_range` |
| `sheet.conditional_formatting` | `sheet.conditional_formats` |
| `value_object[:key]`, `value_object.dig(...)` on page setup, margins, properties, and other metadata | `value_object.key`, or `value_object.to_h[:key]` |
| `xxx` boolean readers such as `row.hidden`, `sheet.fit_to_page` | `xxx?` predicates (`row.hidden?`, `sheet.fit_to_page?`) |

Date cells: `cell.value` returns the Excel serial number. Use `row.values(type_cast: true)`, or `cell.to_date` and `cell.to_time` when `cell.date_format?` is true.

## Writing

| 0.1.x | 0.2.0 |
| :--- | :--- |
| `add_row(values)` | `row(values)` |
| `row(values, styles: [...], style: :x)` | `row(values, cell_styles: [...], row_style: :x)` |
| `set_default_font(...)` | `default_font(...)` |
| `active_tab`, `active_tab = i` | `active_sheet`, `active_sheet = name_or_index` |
| `protect_sheet(...)` | `sheet_protection(...)` or `sheet_protection = {...}` |
| `protect_workbook(...)` | `workbook_protection(...)` or `workbook_protection = {...}` |
| `validate_data(...)` with `col:`, `in: [...]` | `data_validation(...)` with `columns:`, `list: [...]` |
| `select_cell(...)` | `selection = { active_cell:, sqref: }` |
| `page_break_row(i)`, `page_break_col(i)` | `row_break(i)`, `column_break(i)` |
| `group_row`, `group_column`, `group_cols`, `group_col` | `group_rows(range)`, `group_columns(range)` |
| `add_comment`, `add_threaded_comment` | `comment`, `threaded_comment` |
| `add_sparkline` | `sparkline_group` |
| `picture`, `background_picture` | `image`, `background_image =` |
| `merge(row:, col_start:, col_end:)` | `merge("A1:C1")` |
| `image(from_col:, to_col:)`, other `col`/`cols` keywords | `from_column:`, `to_column:`, `column:`, `columns:` |
| `sheet_format(default_col_width:)` | `sheet_format(default_column_width:)` |
| `StyleBuilder#size`, `sz`, `num_fmt`, `num_fmt_id` | `font_size`, `number_format` |
| `StyleBuilder#bold` (and other boolean getters) | `bold?` |
| `Xlsxrb::Utils.col_name_to_index("C")`, `col_index_to_name(2)` | `Xlsxrb.column_index("C")`, `Xlsxrb.column_letter(2)` |
| `Xlsxrb::Utils.ref_to_row_col("C5")`, `row_col_to_ref(4, 2)` | `Xlsxrb.parse_ref("C5")`, `Xlsxrb.cell_to_ref(4, 2)` |
| `Xlsxrb::Utils.date_to_serial`, `time_to_serial`, `serial_to_date`, `serial_to_time` | the same names on `Xlsxrb` |
| `Xlsxrb::MAX_COL` | `Xlsxrb::MAX_COLUMN` |

## Changed behavior

- ZIP compression level 1 is the default (`Xlsxrb::DEFAULT_COMPRESSION_LEVEL`). Writing is faster and files are 5 to 41% larger; pass `compression_level: 6` to `Xlsxrb.write` for the previous size. See [ADR 0009](adr/0009-default-zip-compression-level.md).
- `style(:name, ...)` raises `ArgumentError` for options it does not know (for example `size:` instead of `font_size:`) and for border styles that are not border style names (for example `"slantedDashDot"` instead of `slantDashDot`), and cells, rows, and columns that name an undefined style raise `ArgumentError`. Both were silently ignored before, so code that "worked" may have been writing unstyled cells.
- `escape_formulas:` is gone from `Xlsxrb.write` and `Xlsxrb.build`. It had no effect: strings are always written as text cells, which Excel does not evaluate.
- `Xlsxrb.read` and `Xlsxrb.load` raise `Errno::ENOENT` for a path that does not exist, instead of `Xlsxrb::ZipError`.
- `Xlsxrb.read` and `Xlsxrb.load` raise `Xlsxrb::ParseError` for a ZIP archive that has no workbook part (for example an `.ods` file renamed to `.xlsx`) instead of returning a workbook without sheets.
- `Elements` value objects raise `ArgumentError` for keywords they do not know, including 0.1.x names such as `col:`, `sz:`, `pattern:`, `data_ref:`, `dt:`, and `done:`. The message names the member to use. Before, such keys were dropped without an error.
- `workbook_protection(password: "...")` now protects the workbook with a SHA-512 hash of the password, as `sheet_protection` does. Before, the password was dropped.
- `Elements::FreezePane` and `Elements::SplitPane` have `top_left_cell` instead of `top_row` and `left_column`.
- `rows:` and `columns:` on `data_validation` and `conditional_format` are 0-based and raise `ArgumentError` beyond the last Excel row or column.
- `Date`, `Time`, and `DateTime` cell values keep their date or date-time number format when the cell has a style without a `number_format`, and `DateTime` keeps its time of day.
- The minimum Ruby version is 3.4 (0.1.18 required 4.0).
