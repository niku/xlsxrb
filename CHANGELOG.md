## [Unreleased]

Upgrading from 0.1.x: see [docs/UPGRADING-0.2.md](docs/UPGRADING-0.2.md).

### Added
- Add `show_sheet_tabs?`, `full_calc_on_load?`, and `refresh_all?` boolean predicates on `Elements::WorkbookProperties`.
- Add `core_properties=` and `app_properties=` assignment setters on `WorkbookBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy`.
- Add `name` reader on `StreamWriter::WorksheetProxy` to mirror `WorksheetBuilder#name`.
- Value object delegators (`formula`, `array_formula`, `shared_formula`, `rich_text`, `hyperlink`, `comment`, `cell_error`) on `WorkbookBuilder`, and (`formula`, `array_formula`, `shared_formula`, `rich_text`, `cell_error`) on `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy`.
- Expose `:style_index` as a canonical `Data.define` member on `Elements::Row`.
- Add `columns` 0-argument reader on `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy` returning an array of `Elements::Column`, and on `TableBuilder` returning an array of `Elements::TableColumn`.
- Add `sheet_protection=` assignment setter on `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` accepting `Elements::SheetProtection`, `Hash`, `Boolean`, or `nil`.
- Add `print_titles=` assignment setter on `WorksheetBuilder`, `WorkbookBuilder`, `StreamWriter`, and `WorksheetProxy` accepting `Elements::PrintTitles`, `Hash`, or `nil`.
- Add value object members for data that the reader already parsed but the value objects dropped: `CoreProperties#language`, `WorkbookProperties#update_links`, `WorkbookProtection#revisions_algorithm_name`, `#revisions_hash_value`, `#revisions_salt_value`, and `#revisions_spin_count`, `PageSetup#page_order`, `#horizontal_dpi`, and `#vertical_dpi`, `TableColumn#data_cell_style`, `Font#vert_align`, `ThreadedComment#provider_id`, and `FilterColumn#blank` and `#custom_filters_and`.
- Add `Elements::FilterColumn.from_spec(column_id, spec)` and `FilterColumn#to_spec` to convert between filter column value objects and the `filter_column` hash format (`type: :filters`, `:custom`, `:dynamic`, `:top10`, `:color_filter`, or `:icon_filter`).

### Changed
- Unify property configuration methods across `WorksheetBuilder`, `WorkbookBuilder`, `StreamWriter`, and `WorksheetProxy` (`sheet_properties`, `sheet_view`, `sheet_format`, `print_options`, `workbook_properties`, `core_properties`, `app_properties`) to pure 0-argument getters returning value objects/hashes and keyword argument mutators returning `void` (`nil`).
- Unify single-value property accessors (`print_area`, `auto_filter`) across `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, and `WorkbookBuilder` to 0-argument getters and dedicated `=` assignment setters.
- Unify boolean 0-argument getters to `?` predicate methods across builders: `TableBuilder` (`#show_row_stripes?`, `#show_column_stripes?`, `#show_first_column?`, `#show_last_column?`, `#totals_row?`, `#auto_filter?`), `DataValidationBuilder` (`#allow_blank?`, `#show_input_message?`, `#show_error_message?`), `ConditionalFormatBuilder` (`#bold?`, `#italic?`, `#stop_if_true?`, `#above_average?`, `#equal_average?`, `#percent?`, `#bottom?`), `ChartBuilder` (`#plot_visible_only?`, `#show_legend_key?`), `ChartBuilder::SeriesBuilder` (`#smooth?`), `WorksheetBuilder`/`StreamWriter`/`WorksheetProxy` (`#fit_to_page?`), and `WorkbookBuilder`/`StreamWriter` (`#show_sheet_tabs?`).
- Normalize structured metadata methods (`page_setup`, `page_margins`, `header_footer`, `sheet_protection`, `freeze_pane`, `split_pane`, `selection`, `sort_state`, `outline_properties`, `print_titles`) across `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, and `WorkbookBuilder` to return their current immutable Value Object (or `nil`/Hash) when invoked with 0 arguments, and return `void` (`nil`) when called with configuration arguments.
- Change `TableBuilder#column` to return `void` (`nil`) instead of an internal column definition hash.
- Separate internal pipelines of `Xlsxrb.read` and `Xlsxrb.load`: `Xlsxrb.read` parses minimal streaming metadata directly yielding `StreamSheet` instances, while `Xlsxrb.load` materializes `Elements::Worksheet` directly without allocating intermediate `StreamSheet` instances or double-validating.
- Make `Elements::Cell` fully immutable by freezing instances upon instantiation in `Cell#initialize` and `Cell.fast_create`.
- Restrict `Elements::Workbook#sheets`, `#sheet`, `#[]`, `#each`, and `#active_sheet` signatures to `Elements::Worksheet` instances only.
- Rename `Elements::Styles#num_fmt` to `number_format_code`.
- Rename `Elements::Dxf` member and keyword `:num_fmt` to `:number_format`.
- Rename `Elements::SheetProperties` member and keyword `:page_setup_pr` to `:page_setup_properties`.
- Unify `StyleBuilder` single-attribute methods to dedicated 0-argument getters (`font_size`, `font_name`, `font_color`, `underline`, `vert_align`, `fill_color`, `align_horizontal`, `align_vertical`, `text_rotation`, `indent`, `number_format`) and assignment setters (`font_size=`, `font_name=`, etc.), and 0-argument boolean predicates (`bold?`, `italic?`, `strike?`, `underline?`, `wrap_text?`, `shrink_to_fit?`, `locked?`, `hidden?`, `superscript?`, `subscript?`) with assignment setters (`bold=`, `italic=`, etc.).
- Return frozen `Elements::IgnoredError` value objects from `WorksheetBuilder#ignored_errors`, `StreamWriter#ignored_errors`, and `StreamWriter::WorksheetProxy#ignored_errors`.
- Encapsulate internal methods by making `StyleBuilder#resolve_color` and `StreamSheet#auto_filter_data` private.
- Update blockless `Xlsxrb.read` to return an `Enumerator[StreamSheet, void]` with managed zip container lifecycle; use `Xlsxrb.load` for in-memory random access.
- Update `WorkbookBuilder#sheet` and `StreamWriter#sheet` parameter signatures with explicit keyword arguments `state: nil, tab_color: nil, code_name: nil`.
- Update `StyleBuilder#border` and `StyleBuilder#alignment` parameter signatures with explicit keyword arguments.
- Synchronize `RubyLsp::Xlsxrb::CompletionListener` completion items with the canonical v0.2.0 API.
- Unify `Xlsxrb.array_formula` signature to remove optional positional `ref` in favor of keyword `ref:`.
- Rename `Elements::FilterColumn` member `:col_id` to `:column_id`.
- Rename `Elements::WorkbookProperties` member `:active_tab` to `:active_sheet`, and rename `Elements::Workbook#active_tab` to `#active_sheet_index`.
- Unify `Elements::Worksheet` constructor parameter `sheet_calc_pr:` to `sheet_calc_properties:` and simplify `#with` to canonical keyword arguments.
- Change the default ZIP deflate level from zlib's default (6) to 1, exposed as `Xlsxrb::DEFAULT_COMPRESSION_LEVEL` and used when `compression_level` is omitted or `nil`. Writes are 17 to 43% faster and files 5 to 41% larger in the measurements of [ADR 0009](docs/adr/0009-default-zip-compression-level.md); pass `compression_level: 6` for the previous output.
- Raise `Errno::ENOENT` (or `Errno::EISDIR`) when `Xlsxrb.read` or `Xlsxrb.load` gets a String that is neither XLSX or encrypted file content nor an existing file, instead of parsing the path itself as ZIP content and raising `ZipError: Invalid magic number`. Byte strings that are not an XLSX archive still raise `ZipError`.
- Raise `ArgumentError` for style options that `style` does not know (for example `bolt: true` or `size:` instead of `font_size:`), with a suggestion for the closest valid key, and for a nested group given together with its flat keys (`font: { size: 14 }, bold: true`). These options were silently ignored.
- Raise `ArgumentError` when a cell, row, or column refers to a style name that was never defined, with a suggestion for the closest defined name. Such cells were silently written unstyled. `:default` and `:none`, style indexes, and styles defined later in an `Xlsxrb.build` block are still accepted.
- Raise `ArgumentError` for keyword arguments that an `Elements` value object does not know, in `new` and `with`, including key names that 0.1.x accepted and ignored (for example `FreezePane.new(col: 1)`, `Font.new(sz: 11)`, or `Hyperlink.new(ref: "A1")`). Unknown keys were silently dropped, so a typo or an old name lost the value without an error. The message names the member to use for an old key name, or suggests the closest member for a typo. An old key name given together with its member (as in hashes read from files, which carry both) is dropped. `Chart`, `ConditionalFormat`, `ConditionalFormatRule`, `PivotTable`, `Shape`, `SortState`, `SparklineGroup`, and `Table` still collect other keys in their `options` member.
- Lower the minimum required Ruby version from 4.0 to 3.4, and replace the Ruby 4.0-only `Array#rfind` in the drawing reader with `reverse_each.find`.

### Removed
- Remove dead private helper `Xlsxrb.write_workbook`.
- Remove legacy `unmapped_data` fallbacks for `:style_index` in `Elements::Row#initialize`, and `:style_index`, `:style`, `:best_fit`, and `:collapsed` in `Elements::Column#initialize`.
- Hide non-`?` boolean member reader methods on all 28 `Elements::MetadataRecord` `Data.define` classes and `Elements::Formula#calculate_always`.
- Remove singular property mutators (`sheet_property`, `workbook_property`, `core_property`, `app_property`) across builders in favor of unified `sheet_properties`, `workbook_properties`, `core_properties`, and `app_properties`.
- Remove dual getter/setter `value = OMITTED` parameter branches and non-`?` boolean getter methods (`bold`, `italic`, `strike`, `wrap_text`, `shrink_to_fit`, `locked`, `hidden`, `superscript`, `subscript`) from `StyleBuilder`.
- Remove `WorkbookBuilder#hyperlink` and `WorkbookBuilder#comment` delegators (use top-level `Xlsxrb.hyperlink` and `Xlsxrb.comment`).
- Remove legacy table key fallbacks (`:total_row_formula`, `:total_formula`, `:total_row`, `:total_row_count`) in `DslHelpers` and writers in favor of `:totals_row_formula`, `:totals_row`, and `:totals_row_count`.
- Remove legacy rule `:size` fallback in `DslHelpers.extract_dxf_from_rule!` (retaining `:font_size`).
- Remove legacy duplicate keys (`from_col`/`to_col`, `col:`, `col_fields:`, `dest_ref:`, `active_tab:`, `cols:`, `col_breaks:`, `si:`) across builders and writers in favor of canonical keys (`from_column`/`to_column`, `column:`, `column_fields:`, `destination:`, `active_sheet_index:`, `columns:`, `column_breaks:`, `shared_index:`).
- Remove in-memory `#close` on `Elements::Workbook` and `Elements::Worksheet`, and remove `#load` on `Elements::Worksheet` (`StreamSheet#close` and `StreamSheet#load` remain intact).
- Remove asymmetric legacy methods on `StreamSheet`: `#sheet_views`, `#first_column_index`, `#last_column_index`, `#legacy_drawing`, and `#ext_lst` in favor of `#sheet_view` and `#column_range`.
- Remove `Xlsxrb._read_workbook` and `Elements::Workbook#_load`.
- Remove deprecated alias `DslHelpers.sheet_name` in favor of `DslHelpers.sanitize_sheet_name`.
- Remove unused private cell serialization methods `build_raw_cell`, `build_row_attrs`, and `self.build_raw_cell_from_value` in `lib/xlsxrb.rb`.
- Remove empty `Xlsxrb::Utils` module and `lib/xlsxrb/utils.rb`.
- Remove positional offset arguments and `col_delta:` keyword from `Xlsxrb.expand_formula`, unifying the signature to `def self.expand_formula(formula, from_ref: nil, to_ref: nil, row_delta: nil, column_delta: nil)`.
- Remove singular `column:` and `row:` keyword arguments from `defined_name` across `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` (unified to `columns:` and `rows:`), and align `DslHelpers.normalize_defined_name_value`.
- Remove `show_symbols:` keyword argument from `outline_properties` across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` (unified to `show_outline_symbols:`).
- Remove non-predicate boolean getters across builders: `TableBuilder` (`#show_row_stripes`, `#show_column_stripes`, `#show_first_column`, `#show_last_column`, `#totals_row`, `#auto_filter`), `DataValidationBuilder` (`#allow_blank`, `#show_input_message`, `#show_error_message`), `ConditionalFormatBuilder` (`#bold`, `#italic`, `#stop_if_true`, `#above_average`, `#equal_average`, `#percent`, `#bottom`), `ChartBuilder` (`#plot_visible_only`, `#show_legend_key`), `ChartBuilder::SeriesBuilder` (`#smooth`), `WorksheetBuilder`/`StreamWriter`/`WorksheetProxy` (`#fit_to_page`), `WorkbookBuilder`/`StreamWriter` (`#show_sheet_tabs`), and non-predicate `#escape_formulas` readers on `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` in favor of their `?` predicate equivalents.
- Remove workbook-level delegation methods and stream-lifecycle helpers (`active_sheet`, `active_sheet=`, `first_sheet`, `first_sheet=`, `show_sheet_tabs`, `show_sheet_tabs=`, `tab_ratio`, `tab_ratio=`, `workbook_property`, `workbook_protection`, `core_property`, `app_property`, `custom_property`, `properties`, `default_font`, `sheet_names`, `cleanup!`, `internal_sheet_setup`) from `StreamWriter::WorksheetProxy` to establish full method symmetry with `WorksheetBuilder`.
- Remove positional argument setters from `ChartBuilder` and `SeriesBuilder` in favor of pure 0-argument getters, dedicated `=` assignment setters, `title(**kwargs)`, keyword dictionary methods (`legend(**kwargs)`, `plot_area(**kwargs)`, etc.), and `series(**kwargs, &block)`.
- Remove `#[]`, `#[]=`, and `:in` fallback from `DataValidationBuilder` and `DslHelpers.normalize_data_validation`.
- Remove `#[]`, `#[]=`, `formulas`, `formulas=`, and fallbacks for `:size` and `:format_id` from `ConditionalFormatBuilder`.
- Remove `number_format_id` and `number_format_id=` from `StyleBuilder` in favor of `number_format` and `number_format=`.
- Remove duplicate and ambiguous keys in `StyleBuilder#font` (retaining `:name`, `:size`, `:color`, `:vert_align`) and top-level `StyleBuilder#apply_options!` (retaining `:font_name`, `:font_size`, `:font_color`, `:vert_align`, `:number_format`).
- Remove all remaining deprecated builder methods and aliases across `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy`:
  - Removed `set_default_font`, `active_tab`, `active_tab=`, `protect_workbook`, `protect_sheet`, `validate_data`, `select_cell`, `page_break_row`, `page_break_col`, `add_row`, `skip_row`, `group_column`, `group_cols`, `group_col`, `group_row`, `outline_pr`, `add_sparkline`, `add_comment`, `add_threaded_comment`, `background_picture`, `background_picture=`, `picture`, `picture=`, `add_protected_range`, `add_cell_watch`, `ignore_error`, `add_ignored_error`, `sheet_calc_pr`, `sheet_calc_pr=`, `calc_pr`, `calc_pr=`, `phonetic_pr`, `phonetic_pr=`.
  - Removed `ChartBuilder#secondary_axis`, `ChartBuilder#secondary_axis=`, and `SeriesBuilder#secondary_axis`.
  - Removed deprecated methods on `StyleBuilder`: `size`, `size=`, `sz`, `sz=`, `num_fmt`, `num_fmt=`, `num_fmt_id`, `num_fmt_id=`.
  - Removed deprecated methods on `TableBuilder`: `add_column`, `total_row`, `total_row=`, `total_row_count`, `total_row_count=`.
  - Removed deprecated methods on `DataValidationBuilder`: `range`, `range=`, `col`, `col=`, `row`, `row=`, `in`, `in=`, `error`, `error=`, `prompt`, `prompt=`.
  - Removed deprecated methods on `ConditionalFormatBuilder`: `range`, `range=`, `col`, `col=`, `row`, `row=`, `format_id`, `format_id=`, `size`, `size=`, `sz`, `sz=`.
  - Removed dual getter/setter argument branches across builders (`active_sheet`, `first_sheet`, `show_sheet_tabs`, `tab_ratio`, `code_name`, `fit_to_page`, `tab_color`, `sheet_calc_properties`, `phonetic_properties`, `background_image`, `state`, `auto_filter`, `print_area`), strictly enforcing zero-argument getters and dedicated arity-1 assignment setters (`xxx=`).
  - Removed deprecated keyword arguments: `styles:` and `style:` on `row` (use `cell_styles:` and `row_style:`); coordinate kwargs `row:`, `col_start:`, `col_end:`, `row_start:`, `row_end:` on `merge` (use `merge(range)`); `from_col:` and `to_col:` on `image` and `shape` (use `from_column:` and `to_column:`); `default_col_width:`, `base_col_width:`, and `outline_level_col:` on `sheet_format` (use `default_column_width:`, `base_column_width:`, `outline_level_column:`); `cols:` on `print_titles` (use `columns:`); `col:`, `ref_sheet:`, and `local:` on `defined_name` (use `column:` and `scope:`); `col:` and `except_col:` on `data_validation`, `conditional_format`, and `ignored_error` (use `columns:` and `except_columns:`); `si:` on `Formula` (use `shared_index:`).
- Remove `StreamWriter#shared_strings` reader method in favor of `shared_strings?`.
- Remove constant `Xlsxrb::MAX_COL` in favor of `Xlsxrb::MAX_COLUMN`.
- Remove `Xlsxrb::Error::DIV0`..`VALUE` cell error constants on the exception class in favor of `Elements::CellError::*`.
- Remove `Elements::RichTextRun#sz` in favor of `font_size`.
- Replace the `top_row` and `left_column` members of `Elements::FreezePane` and `Elements::SplitPane` with `top_left_cell`. Neither member was read or written.
- Remove the `escape_formulas:` option of `Xlsxrb.write`, `Xlsxrb.build`, `StreamWriter`, `WorkbookBuilder`, and `WorksheetBuilder`, with its `escape_formulas?` and `escape_formulas=` accessors. It was passed down to the worksheet writer but never used: output was identical with and without it. Strings are always written as text cells, which Excel does not evaluate as formulas, with or without the option.
- Remove deprecated `Xlsxrb.modify`, `Elements::Workbook#update_sheet`, and `Elements::Worksheet#update_cell` in favor of streaming transformations (`Xlsxrb.read` + `Xlsxrb.write`) or template generation via `Xlsxrb.write`.
- Remove Hash-like access (`#[]`, `#key?`, `#dig`, `#fetch`, `#merge`, and `== Hash`) on `Elements::MetadataRecord` value objects and `Elements::Hyperlink#[]` in favor of direct reader method calls and `#to_h`.
- Remove deprecated methods on `Elements::Row`: `pad_empty_cells` (use `padded`), `cells_hash` (use `cells`), `formulas` (use `formula_expressions`), and bare boolean queries (`hidden`, `custom_height`, `collapsed` in favor of `?`-suffixed predicates).
- Remove deprecated methods on `StreamRow`: `to_h` (use `values` or `cells`), `formulas` (use `formula_expressions`), and bare boolean queries (`hidden`, `custom_height`, `collapsed` in favor of `?`-suffixed predicates).
- Remove deprecated methods on `Elements::Cell`: `content` (use `value`), `column_index`, `column_letter`, `parse_ref`, `date1904` (use `date1904?`), `link?` (use `hyperlink?`), and `threaded_comments?` (use `threaded_comment?`).
- Remove keyword argument `date1904:` on `Cell#to_date` and `Cell#to_time`.
- Remove deprecated `StreamSheet#cell` and `Elements::Worksheet#cell` in favor of coordinate bracket access (`sheet[...]`).
- Remove deprecated `StreamSheet#first_row` and `StreamSheet#last_row` in favor of `row_range` or `dimension_bounds`.
- Remove deprecated `CoordinateAccess#cells_hash`, `first_column`, `first_col`, `last_column`, `last_col`, and `cell_value`.
- Remove deprecated `Elements::Image#cell` and `cell_ref` (use `ref`), `from_col` (use `from_column`), and `to_col` (use `to_column`).
- Remove remaining deprecated aliases, duplicate Data members, and initialization fallbacks across `Elements::*` Value Objects:
  - Remove duplicate `Data.define` members: `:cols` from `Elements::PrintTitles` (use `:columns`), `:col` from `Elements::FreezePane` (use `:column`), `:first_tab` from `Elements::WorkbookProperties` (use `:first_sheet`), `:outline_pr` from `Elements::SheetProperties` (use `:outline_properties`), `:rel_id`, `:ext`, and `:filename` from `Elements::BackgroundImage` (use `:r_id`, `:extension`, `:path`), and `:location` from `Elements::PivotTable` (use `:ref`).
  - Remove `Elements::RichTextRun#[]` and `other.is_a?(Hash)` branches from `Elements::Hyperlink#==`, `Elements::CustomProperty#==`, and `Elements::Styles#==`.
  - Remove non-predicate formula aliases on `Elements::Formula`: `always_calculate_array`, `assign_to_name`, `data_table_2d`, and `data_table_row` in favor of predicate methods (`?`).
  - Remove `Elements::Column#letter` in favor of canonical `column_letter`.
  - Remove `Elements::Styles#format_code` in favor of `number_format`.
  - Remove `StreamSheet#_load` in favor of `StreamSheet#load`.
  - Remove deprecated `conditional_formatting:` and `col_breaks:` keyword arguments in `Elements::Worksheet#initialize` in favor of `conditional_formats:` and `column_breaks:`.
  - Remove legacy keyword arguments and aliases on element Value Objects: `SheetFormat` (`default_col_width`, `outline_level_col`, `base_col_width`), `FilterColumn` (`hidden_button:`), `PageBreak` (`manual:`), `Comment` (`cell:` and `alias cell ref`), `ThreadedComment` (`cell:`, `person_id:`, `dt:`, `timestamp:`, `done:`, and `alias done resolved`), `DataValidation` (`show_dropdown` and `show_dropdown?`), `ConditionalFormatRule` (`format_id` and `alias formulas formula`), `TableColumn` (`data_formula:` and `alias data_formula calculated_column_formula`), `Sparkline` (`location_ref`, `location`, `data_ref`, `f`, `data`), `SparklineGroup` (`display_empty`, `max`, `min`), `Chart` (`rid`, `chart_type`, `from_col`, `to_col`), `HeaderFooterImage` (`rid`, `ext`, `filename`), `Font` (`sz`), `Fill` (`pattern`), and `CellXf` (`number_format_id`).
- Remove deprecated `Elements::Worksheet#dimension=` (use `with(dimension: ...)`), `Elements::Worksheet#with_images`, and `StreamSheet#with_images`.
- Remove deprecated `Elements::Worksheet#conditional_formatting` and `StreamSheet#conditional_formatting` in favor of `conditional_formats`.
- Remove deprecated `Elements::Styles#add_xf`, `add_num_fmt`, and `num_fmts` in favor of `add_cell_xf`, `add_number_format`, and `number_formats`.
- Remove deprecated `Elements::Styles#xf`, `num_fmt`, `format_code_for`, `classification_for`, and `sz`.
- Remove deprecated `Xlsxrb::Utils` coordinate and date conversion methods (`col_name_to_index`, `col_index_to_name`, `ref_to_row_col`, `row_col_to_ref`, `parse_range`, `range_to_ref`, `date_to_serial`, `serial_to_date`, `serial_to_time`, `time_to_serial`, `split_coordinate`, `valid_excel_date?`, `expand_formula`, `datetime_to_serial`) in favor of canonical `Xlsxrb.*` entry points.
- Remove deprecated transitional top-level aliases on `Xlsxrb`: `datetime_to_serial`, `serial_to_datetime`, `color_to_hex`, `sheet_name`, `stream`, and `format` (use `Xlsxrb.format_value`).
- Remove deprecated transitional aliases `to_workbook` on `Workbook`, `to_worksheet` on `Worksheet` and `StreamSheet`, `each_sheet` on `Workbook`, and `col_breaks` on `Worksheet` and `StreamSheet`.
- Remove non-predicate boolean queries on element readers (`trim_empty_rows`, `pad_empty_rows`, `pad_empty_cells`, `fit_to_page`, `show_sheet_tabs`, `full_calc_on_load`) in favor of `?`-suffixed methods (`trim_empty_rows?`, `pad_empty_rows?`, `pad_empty_cells?`, `fit_to_page?`, `show_sheet_tabs?`, `full_calc_on_load?`).

### Fixed
- Protect the workbook with `workbook_protection(password:)`. `Elements::WorkbookProtection` had no `password` member, so the password was dropped and the workbook was written without one. A plain-text password is now hashed with SHA-512, as for `sheet_protection`, and a 4-digit hexadecimal value is written as the legacy `workbookPassword`. `Xlsxrb.load` also keeps the hash, revision lock, and revisions password attributes of `<workbookProtection>`, which it dropped, so that `Xlsxrb.write` writes them back.
- Keep filter column conditions in `Elements::Worksheet#filter_columns` and `StreamSheet#filter_columns`. Hashes in the `filter_column` format (`type:`, `values:`, `operator:`, ...) were passed to `Elements::FilterColumn.new`, which dropped those keys, so the readers returned empty filter columns and `Xlsxrb.write` of value objects wrote an `<autoFilter>` without its conditions.
- Write threaded comments built from `Elements::ThreadedComment` with their `created_at` time and `resolved` state; the writer only looked at the `dt`/`timestamp` and `done` keys.
- Write the `horizontalDpi` and `verticalDpi` attributes of page setup when they are positive; the schema requires at least 1.
- Keep the pattern type of differential format fills loaded from `styles.xml` (`Elements::Dxf#fill`). The styles reader names it `pattern:`, which `Elements::Fill` dropped.
- Wrap keyword arguments with double splat in Data.define constructors across Elements to prevent heap corruption and segfaults in CRuby 4.0 Ractors.
- Keep the top-left cell of freeze and split panes. `freeze_pane(top_left_cell:)` and the reader both used a `top_left_cell` that `Elements::FreezePane` and `Elements::SplitPane` did not have, so the value objects dropped it and the writer always wrote the cell just past the frozen rows and columns; the reader also did not read `topLeftCell` for frozen panes.
- Keep the time of day of `DateTime` values. `Xlsxrb.date_to_serial`, `Xlsxrb.time_to_serial`, and cell values written by `Xlsxrb.write` and `Xlsxrb.build` treated `DateTime` (a subclass of `Date`) as a date, dropping the time and applying the date-only format. They now handle it like `Time`.
- Write worksheets whose `selection` is an Array, one selection per pane, as `Elements::Worksheet` accepts. `Xlsxrb.write` called `to_h` on the Array and raised `TypeError`.
- Write workbooks loaded from files with rich-text comments, such as comments that Excel starts with the author's name in bold. `MetadataRecord#to_h` called `Enumerable#to_h` on the `Elements::RichText` text of `Elements::Comment`, so `Xlsxrb.write` raised `TypeError`. `to_h` now keeps `RichText` values as they are.
- Write workbooks loaded from files whose `docProps/app.xml` lists `HeadingPairs`, as files saved by Excel do. `MetadataRecord#to_h` called `to_h` on the `[name, count]` pairs of `AppProperties#heading_pairs`, so `Xlsxrb.write` raised `TypeError`.
- Keep the date or time number format on `Date` and `Time` cell values that are given a cell style without a `number_format` (for example `cell_styles: [:bold]`). The style replaced the automatic `yyyy-mm-dd` format, so Excel showed the serial number. Such cells now get a variant of their style that adds the default format, in both `Xlsxrb.write` blocks and `Xlsxrb.build`.
- Inherit `sheet_format` `default_column_width` as the width of columns that define `style:` without `width:`. The column writer still looked up the removed `default_col_width` key and dropped the width.
- Raise `ArgumentError` when `rows:` or `columns:` on `data_validation` and `conditional_format` reach outside the Excel grid (for example `rows: 1..Xlsxrb::MAX_ROW`, which is one past the last 0-based row) instead of writing an invalid reference such as `A2:A1048577`.
- Keep the REXML entity expansion limits that the application set before `require "xlsxrb"`. Loading xlsxrb reset `REXML::Security.entity_expansion_limit` and `entity_expansion_text_limit` to the REXML defaults, which undid hardening such as an application setting the limit to 0 to disable entity expansion.
- Keep reading entries after one that ends with a data descriptor when `ZipReader` falls back to walking local headers (an archive without a readable central directory). Ruby's `Zlib::Inflate` drops input past the end of the deflate stream, so the reader skipped over the following entries; it now rewinds by the compressed length, and reads 8-byte sizes from the data descriptor of ZIP64 entries.
- Build cells in streaming and in-memory reads with one direct initializer instead of an `instance_variable_set` call per attribute, which had grown from 7 to 14 calls per cell since 0.1.10. A 1,000,000-cell streaming read is about 12% faster.
- Take the `r`, `s`, and `t` cell attributes from the cell scanning regex when they appear in that order, as Excel and xlsxrb write them, instead of three more regex lookups per cell (other orders still fall back to those lookups), and stop copying cell values that need no whitespace stripping or re-encoding. A 1,000,000-cell streaming read is about 16% faster and allocates 43% fewer objects.
- Check cell reference digits with plain byte comparisons instead of `Comparable#between?`, and yield cells instead of calling a block Proc, in the cell scanning loops (about 3% faster streaming reads).
- Consume inflated ZIP entry data by offset instead of `slice!`, which moved the rest of each inflated chunk on every read, and stop re-encoding chunks that are already binary. Streaming reads run half as many garbage collections.
- Write ZIP64 records when an entry size, a file offset, or the entry count exceeds the 32-bit (16-bit for counts) ZIP fields. Sizes were truncated to 32 bits, so a worksheet whose XML exceeded 4 GiB uncompressed produced a corrupt XLSX without an error. Archives below these limits are written unchanged, except that entries streamed to unseekable output are declared ZIP64 up front, because their size is unknown when the local header is written. `ZipReader` also reads the ZIP64 end of central directory record. See [ADR 0008](docs/adr/0008-zip64-only-when-a-limit-is-exceeded.md).

## [0.1.18] - 2026-11-07

### Added
- Add Value Object factory delegator methods (`formula`, `array_formula`, `shared_formula`, `rich_text`, `comment`, `cell_error`, `hyperlink`) across `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy`.
- Add canonical coordinate conversion facade methods `Xlsxrb.parse_ref`, `Xlsxrb.cell_to_ref`, `Xlsxrb.column_index`, and `Xlsxrb.column_letter`.
- Add `Xlsxrb.format_value` as the canonical value formatting method.
- Add `Xlsxrb.time_to_serial` (and `Utils.time_to_serial`) for converting Time/Date objects to Excel numeric serial numbers.
- Add `Xlsxrb::MAX_COLUMN` constant (16,384) representing the maximum column count supported by Excel / OpenXML.
- Support passing nested bounding coordinate arrays `[[top_row, left_col], [bottom_row, right_col]]` (as returned by `parse_range`) to `Xlsxrb.range_to_ref`.
- Support `column_delta:` keyword in `Xlsxrb::Utils.expand_formula`.
- Support `ref:` keyword argument in `Xlsxrb.array_formula` and include `Elements::CellError` in `cached_value:` types.
- Support `new_password:` and writer/reader option forwarding in `Xlsxrb.modify`, and return the block result.
- Define Ractor-shareable frozen `Elements::CellError` constants (`DIV0`, `NA`, `NAME`, `NULL`, `NUM`, `REF`, `VALUE`, `GETTING_DATA`) on `Elements::CellError`.
- Add canonical cell-level and row-level style keywords `cell_styles:` and `row_style:` across `WorksheetBuilder#row`, `StreamWriter#row`, and `StreamWriter::WorksheetProxy#row`.
- Expose `:style_index` as a canonical `Data.define` member on `Elements::Row`.
- Add `style:` keyword to `Elements::Cell.new` and `Elements::Worksheet#update_cell` for style names and inline style definitions.
- Add 1-based `row_range` and `column_range` boundaries to `Elements::Worksheet` (via `CoordinateAccess`) and `StreamSheet` (from sheet dimension bounds), returning idiomatic `Range[Integer]` boundaries.
- Support 2D coordinate indexing `sheet[row, col]` and letter column indexing `sheet[row, "A"]` on `Elements::Worksheet` and `CoordinateAccess`.
- Add `column` and `columns` builder methods across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` with unified options.
- Support 0-based page break indices (`row_breaks`, `column_breaks`) across builders and readers.
- Add `Elements::WorkbookProperties#active_sheet` and `Elements::Workbook#active_sheet_index`.
- Add `Elements::FilterColumn#column_id`.
- Add structured Hash access for `category_axis` and `value_axis` titles on `Elements::Chart`.
- Add dedicated assignment setters returning assigned value (`active_sheet=`, `first_sheet=`, `show_sheet_tabs=`, `tab_ratio=`, `code_name=`, `fit_to_page=`, `tab_color=`, `phonetic_properties=`, `sheet_calc_properties=`, `background_image=`, `totals_row=`, `totals_row_count=`, `state=`, `print_area=`, `print_titles=`) across `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::WorksheetProxy`.
- Add zero-argument property getters across builders and `StyleBuilder`.
- Add predicate methods (`bold?`, `italic?`, `strike?`, `wrap_text?`, `shrink_to_fit?`, `locked?`, `hidden?`, `superscript?`, `subscript?`) and assignment setters (`font=`, `fill=`, `border=`, `vert_align=`, `superscript=`, `subscript=`, `wrap_text=`, `shrink_to_fit=`, `locked=`, `hidden=`, `protection=`, `number_format=`) on `StyleBuilder`.
- Re-affirm `totals_row`, `totals_row=`, `totals_row_count`, and `totals_row_count=` as canonical on `TableBuilder` to align with ECMA-376 `totalsRowCount`.
- Promote reader metadata to immutable Data value objects across `Elements::Styles`, `Elements::Workbook`, and `Elements::Cell`.
- Add `Elements::Styles#number_format_code(id)`.
- Support Symbol sheet lookup in `Elements::Workbook#sheet` and `Elements::Workbook#[]`.
- Freeze `Elements::Cell` instances on instantiation (`initialize` and `fast_create`).
- Return frozen `Array[Elements::IgnoredError]` from `WorksheetBuilder#ignored_errors`, `StreamWriter#ignored_errors`, and `WorksheetProxy#ignored_errors`.
- Add XML element existence guards for `merged_cells`, `data_validations`, `conditional_formats` on `StreamSheet`.
- Establish four-pillar I/O architecture: `Xlsxrb.read`, `Xlsxrb.write`, `Xlsxrb.load`, and `Xlsxrb.build`.
- Direct worksheet materialization in `Xlsxrb.load` without intermediate `StreamSheet` allocations.
- Support on-demand zip reopen after close in `Ooxml::ZipReader`.
- Clarify `@api private` on `Ooxml::Reader`, `Ooxml::Writer`, and `Ooxml::ZipGenerator`.

### Changed
- Update `Elements::Row#[]` and `StreamRow#[]` to resolve column indices and column letters via `cell_at` instead of indexing the raw sparse `cells` array, returning `nil` when a cell at the specified column is absent.
- Separate builder getter (`xxx`) and setter (`xxx=`) methods, allowing explicit `nil` resets.
- Encapsulate `StyleBuilder#resolve_color` as `private`.
- Encapsulate `StreamSheet#auto_filter_data` as `private`.
- Make `Xlsxrb._read_workbook` `private_class_method`.

### Deprecated
- `DslHelpers.sheet_name` is deprecated in favor of `DslHelpers.sanitize_sheet_name`; will be removed in version 0.2.0.
- Positional `pos_ref` in `Xlsxrb.array_formula` is deprecated in favor of keyword `ref:`; will be removed in version 0.2.0.
- Top-level aliases `Xlsxrb.stream`, `Xlsxrb.sheet_name`, `Xlsxrb.color_to_hex`, `Xlsxrb.serial_to_datetime`, and `Xlsxrb::Utils.serial_to_datetime` are deprecated in favor of `Xlsxrb.read`, `Xlsxrb.sanitize_sheet_name`, `Xlsxrb.to_hex_color`, and `Xlsxrb.serial_to_time`; will be removed in version 0.2.0.
- `Xlsxrb.format` is deprecated in favor of `Xlsxrb.format_value`; will be removed in version 0.2.0.
- `Xlsxrb.datetime_to_serial` and `Xlsxrb::Utils.datetime_to_serial` are deprecated in favor of `Xlsxrb.time_to_serial`; will be removed in version 0.2.0.
- `Xlsxrb::MAX_COL` is deprecated in favor of `Xlsxrb::MAX_COLUMN`; will be removed in version 0.2.0.
- `Xlsxrb::Utils` coordinate and date conversion methods (`col_name_to_index`, `col_index_to_name`, `ref_to_row_col`, `row_col_to_ref`, `parse_range`, `range_to_ref`, `date_to_serial`, `serial_to_date`, `serial_to_time`, `time_to_serial`) and `Elements::Cell.column_index`, `column_letter`, `parse_ref` are deprecated in favor of top-level `Xlsxrb.*` helpers; will be removed in version 0.2.0.
- Keyword `:col_delta` in `Xlsxrb::Utils.expand_formula` is deprecated in favor of `:column_delta`; will be removed in version 0.2.0.
- Keyword `:si` in `Xlsxrb.shared_formula` is deprecated in favor of `:shared_index`; will be removed in version 0.2.0.
- `Xlsxrb::Error::DIV0`..`VALUE` cell error constants on the exception class are deprecated in favor of `Elements::CellError::DIV0`..`VALUE`; will be removed in version 0.2.0.
- Redundant method and setter aliases across builders and element readers (`add_row`, `skip_row`, `group_column`, `group_cols`, `group_col`, `group_row`, `outline_pr`, `add_threaded_comment`, `background_picture`, `picture`, `background_picture=`, `picture=`, `add_protected_range`, `add_cell_watch`, `phonetic_pr`, `phonetic_pr=`, `set_default_font`, `add_column`, `first_col`, `last_col`, `range`, `range=`, `col`, `col=`, `row`, `row=`, `sz`, `sz=`, `error`, `error=`, `prompt`, `prompt=`) are deprecated and will be removed in version 0.2.0.
- Singular `column:` and `row:` keywords in `defined_name` are deprecated in favor of plural `columns:` and `rows:`; will be removed in version 0.2.0.
- `show_symbols:` keyword in `outline_properties` is deprecated in favor of `show_outline_symbols:`; will be removed in version 0.2.0.
- `fit_to_page?` on builders is deprecated in favor of `fit_to_page`; will be removed in version 0.2.0.
- Keywords `styles:` and `style:` on `WorksheetBuilder#row`, `StreamWriter#row`, and `WorksheetProxy#row` are deprecated in favor of `cell_styles:` and `row_style:`; will be removed in version 0.2.0.
- Passing non-Integer style names or Hash definitions to `style_index:` on `Elements::Cell.new` and `Elements::Worksheet#update_cell` is deprecated in favor of `style:`; will be removed in version 0.2.0.
- `Elements::Worksheet#cell` is deprecated in favor of bracket access (`sheet["A1"]` or `sheet[row, col]`) and will be removed in version 0.2.0.
- `StreamSheet#cell` is deprecated in favor of materializing an `Elements::Worksheet` via `sheet.load` and using coordinate bracket access; will be removed in version 0.2.0.
- `StreamSheet#first_row` and `StreamSheet#last_row` returning integer row numbers are deprecated in favor of `row_range` or `dimension_bounds`; will be removed in version 0.2.0.
- Passing coordinate keyword arguments (`row:`, `col_start:`, `col_end:`, `row_start:`, `row_end:`) to `merge` on `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` is deprecated in favor of coordinate range strings (e.g. `"A1:C1"`); coordinate kwargs will be removed in version 0.2.0.
- `Elements::FilterColumn#col_id` is deprecated in favor of `#column_id`; will be removed in version 0.2.0.
- `Elements::WorkbookProperties#active_tab` is deprecated in favor of `#active_sheet`, and `Elements::Workbook#active_tab` is deprecated in favor of `#active_sheet_index`; will be removed in version 0.2.0.
- `Elements::Chart#cat_axis_title` and `#val_axis_title` are deprecated in favor of `category_axis[:title]` and `value_axis[:title]`; will be removed in version 0.2.0.
- `Elements::Worksheet` constructor keywords `sheet_calc_pr:`, `calc_pr:`, `phonetic_pr:`, `background_picture:`, `picture:` are deprecated in favor of canonical keyword arguments; will be removed in version 0.2.0.
- Calling builder getter methods with arguments to set values or mutate state (`active_sheet`, `first_sheet`, `show_sheet_tabs`, `tab_ratio`, `code_name`, `tab_color`, `sheet_calc_properties`, `phonetic_properties`, `background_image`, `state`, and subbuilders `TableBuilder`, `DataValidationBuilder`, `ConditionalFormatBuilder`) is deprecated and will become 0-argument getters in version 0.2.0; use dedicated setters (`xxx=`) instead.
- Calling `fit_to_page` with arguments is deprecated and will become a 0-argument boolean getter in version 0.2.0; use `fit_to_page = true/false` or `page_setup(fit_to_width: ..., fit_to_height: ...)` instead.
- Calling `StyleBuilder` boolean mutator methods (`bold`, `italic`, `underline`, `strike`, `superscript`, `subscript`, `wrap_text`, `shrink_to_fit`, `locked`, `hidden`) with zero arguments to mutate state is deprecated; pass `true` explicitly (e.g. `bold(true)`).
- `TableBuilder#[]` and `#[]=` are deprecated in favor of `#column(name)` and `#add_column`; will be removed in version 0.2.0.
- `TableBuilder#total_row`, `#total_row=`, `#total_row_count`, and `#total_row_count=` are deprecated in favor of `totals_row`, `totals_row=`, `totals_row_count`, and `totals_row_count=`; will be removed in version 0.2.0.
- `Elements::Styles#num_fmt` is deprecated in favor of `Styles#number_format_code`; will be removed in version 0.2.0.
- `Elements::Dxf#num_fmt` is deprecated in favor of `Dxf#number_format`; will be removed in version 0.2.0.
- `Elements::SheetProperties#page_setup_pr` is deprecated in favor of `SheetProperties#page_setup_properties`; will be removed in version 0.2.0.
- `StreamSheet#first_column` and `StreamSheet#last_column` are deprecated in favor of `column_range`; will be removed in version 0.2.0.
- `StreamSheet#sheet_views`, `#first_column_index`, `#last_column_index`, `#legacy_drawing`, and `#ext_lst` are deprecated in favor of `#sheet_view` and `#column_range`; will be removed in version 0.2.0.
- `Elements::Worksheet#conditional_formatting` and `StreamSheet#conditional_formatting` are deprecated in favor of `conditional_formats`; will be removed in version 0.2.0.
- `Elements::Row#to_a` and `StreamRow#to_a` returning cell values are deprecated in favor of `row.values`; in version 0.2.0, `to_a` will return `Array[Elements::Cell]` to conform to `Enumerable#to_a`.
- Non-predicate boolean queries on element readers (`trim_empty_rows`, `pad_empty_rows`, `pad_empty_cells`, `fit_to_page`, `show_sheet_tabs`, `full_calc_on_load`) are deprecated in favor of `?`-suffixed methods (`trim_empty_rows?`, `pad_empty_rows?`, `pad_empty_cells?`, `fit_to_page?`, `show_sheet_tabs?`, `full_calc_on_load?`); will be removed in version 0.2.0.
- Transitional aliases `to_workbook` on `Workbook`, `to_worksheet` on `Worksheet` and `StreamSheet`, and `each_sheet` on `Workbook` are deprecated in favor of `load` and `each`; will be removed in version 0.2.0.
- Style aliases `Styles#xf`, `Styles#format_code_for`, `Styles#classification_for`, and `StyleBuilder#num_fmt` are deprecated in favor of `cell_xf`, `number_format`, and `format_type`; will be removed in version 0.2.0.
- `Xlsxrb.modify` is deprecated in favor of explicit `Xlsxrb.load` -> mutate -> `Xlsxrb.write` pipeline; will be removed in version 0.2.0.
- `Xlsxrb.write_workbook` and `Xlsxrb.workbook_to_data` are deprecated in favor of `Xlsxrb.write`; will be removed in version 0.2.0.
- In-memory `Elements::Workbook#close`, `Elements::Worksheet#close`, and `Elements::Worksheet#load` are deprecated; will be removed in version 0.2.0 (`StreamSheet#close` and `StreamSheet#load` remain supported).
- Returning non-`Elements::Workbook` values from `Xlsxrb.modify` blocks is deprecated and will raise `TypeError` in version 0.2.0.

### Fixed
- Prevent attribute loss in `Elements::PivotTable`, `Elements::TableColumn`, `Elements::ThreadedComment`, `Elements::Comment`, and `Elements::FilterColumn` by retaining full schema attributes and capturing unmapped keyword options into `options:`.
- Eliminate internal calls to deprecated `MetadataRecord#[]`, `#merge`, and `cols:` across `WorksheetBuilder`, `WorkbookBuilder`, `StreamSheet`, `Elements::Worksheet`, and `Elements::Styles`.

## [0.1.17] - 2026-11-01

### Added
- Support threaded comments (`<ThreadedComments>`) and workbook person lists (`<personList>`) across `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, `Ooxml::Writer`, `WorkbookWriter`, `DrawingXml`, `Elements::Worksheet`, `Elements::Cell`, `StreamSheet`, and `Ooxml::Reader`, serialize `xl/threadedComments/threadedComment{N}.xml` and `xl/persons/person.xml` conforming to the OpenXML threaded comments specification, support root comments and replies with `parent_id`, timestamps, and resolution states (`done`), register relationship types and content types, and provide query accessors across worksheets and cells.
- Support combo charts and secondary value axes across `ChartBuilder`, `DrawingXml`, `Ooxml::Writer`, and `DrawingListeners`, allow mixed chart types (e.g. bar and line) inside a single plot area with primary and secondary value axes (`secondary_value_axis:`, `secondary_axis:`), support per-series axis binding via `SeriesBuilder#secondary_axis` and `SeriesBuilder#axis`, serialize distinct `<c:valAx>` elements with independent scaling, orientation, units, and gridlines conforming to ECMA-376 Part 4 §21.2 schema sequence, and expose secondary value axis configurations in `Ooxml::Reader`.
- Support shared formula (`<f t="shared">`) serialization and reader expansion across `Xlsxrb.shared_formula`, `Xlsxrb.formula(..., shared_index:, si:)`, `WorksheetWriter`, `FeaturesXml`, `WorksheetParser`, and `WorksheetListener`, serialize master (`ref`, `si`) and follower (`si`) elements conforming to ECMA-376 Part 4 §18.3.1.40, and add `Xlsxrb::Utils.expand_formula` to expand relative cell references across shared formula ranges during SAX and streaming parsing.
- Support table calculated columns (`calculatedColumnFormula`) and custom totals row formulas (`totalsRowFormula`) across `TableBuilder`, `DslHelpers.normalize_table_options`, `WorkbookWriter`, `Writer`, and `FeatureListeners`, add `TableBuilder#calculated_column` helper, support `formula:`, `data_formula:`, `totals_row_formula:`, `total_row_formula:`, and `total_formula:` aliases, automatically infer `totalsRowFunction="custom"` when a totals formula is provided without an explicit function, and expose `calculated_column_formula` and `data_formula` in reader output.
- Support font text decoration attributes, including underline styles (`:single`, `:double`, `:single_accounting`, `:double_accounting`, `:none`) and vertical alignment (`:superscript`, `:subscript`, `:baseline`), with dedicated builder helpers (`superscript`, `subscript`, `vert_align`, `underline`) and option mappings across `StyleBuilder`, `WorkbookWriter`, `StylesXml`, `FeaturesXml`, `WorksheetWriter`, and `Ooxml::Reader`.
- Support row and column grouping DSL (`group_rows`, `group_columns`, `group_row`, `group_column`, `group_cols`, `group_col`) and outline view properties (`outline_properties`, `outline_pr`) across `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, `Elements::Worksheet`, `StreamSheet`, and `Ooxml::Reader`, emit schema-compliant `<outlinePr applyStyles="..." summaryBelow="..." summaryRight="..." showOutlineSymbols="..."/>` in `WorksheetWriter`, and emit `collapsed="1"` on `<col>` elements.
- Enhance conditional formatting with statistical and relative rules (`above_average`, `top10`, `duplicate_values`, `unique_values`, `contains_blanks`, `not_contains_blanks`, `contains_errors`, `not_contains_errors`, `time_period`), add builder accessors (`above_average`, `equal_average`, `std_dev`, `rank`, `percent`, `bottom`, `text`, `time_period`) to `ConditionalFormatBuilder`, map symbol time periods (`:today`, `:last_7_days`, `:this_month`, etc.) to ECMA-376 `ST_TimePeriod` values, and automatically infer rule type from options.
- Add `ConditionalFormatBuilder` supporting block configuration for conditional formatting across `WorksheetBuilder`, `StreamWriter`, `WorksheetProxy`, and `Ooxml::Writer`.
- Add coordinate targeting (`columns:`, `rows:`, `col:`, `row:`, `except_rows:`) to `conditional_format` and `ConditionalFormatBuilder`, symmetrically resolving to `sqref`.
- Enhance `color_scale` conditional formatting with automatic cfvo generation for 2-color (`min`/`max`) and 3-color (`min`/50%/`max`) scales, and auto-normalize 6-digit RGB hex colors to 8-digit ARGB.
- Enhance `icon_set` conditional formatting with automatic cfvo percent distribution for 3, 4, and 5 icon sets, and support symbol icon set styles (`:three_arrows`, `:four_rating`, `:five_quarters`, etc.).
- Enhance `table` and `TableBuilder` with `auto_filter:` option, boolean `total_row` / `totals_row` helpers, symbol column names, and symbol `totals_row_function` mappings (`:sum`, `:average`, `:count`, `:min`, `:max`, `:count_nums`, `:std_dev`, `:var`, `:custom`).
- Add worksheet name sanitization utility via `Xlsxrb.sanitize_sheet_name` (aliased as `Xlsxrb.sheet_name` and `DslHelpers.sanitize_sheet_name`), sanitizing forbidden characters (`\`, `/`, `?`, `*`, `[`, `]`, `:`, control characters), replacing leading/trailing apostrophes, appending `replace_char` to the reserved `History` sheet name, and truncating names to 31 characters.
- Add `sheet_names` to `StreamWriter`, `WorksheetProxy`, and `WorkbookBuilder` returning an array of worksheet names currently registered in the workbook.
- Support `sheet.defined_name` symmetrically across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, automatically quoting sheet names with single quote escaping and generating absolute `$` coordinate references.
- Support `fill: { color: ... }` and `fill_color:` in conditional formatting rules and `extract_dxf_from_rule!`.
- Support `size:` alongside `sz:` in conditional formatting differential formatting (`dxf`) fonts.
- Support `base_font:` in `Xlsxrb.rich_text(*runs, base_font: ...)` to merge base font attributes across text runs.
- Support `size:` alongside `sz:` in `RichTextRun`, inline rich text, and shared strings XML fonts.
- Add coordinate-based targeting (`columns:`, `rows:`, `except_rows:`, `col:`, `row:`) to `validate_data` and `DataValidationBuilder` with interval run compression (`compress_runs`), complement calculation (`complement_runs`), and automatic chunking into multiple `<dataValidation>` entries when reference counts exceed 32,767 (`Xlsxrb::MAX_REFS_PER_VALIDATION`).
- Expose spreadsheet dimension limits `Xlsxrb::MAX_ROW = 1_048_576`, `Xlsxrb::MAX_COL = 16_384`, and `Xlsxrb::MAX_REFS_PER_VALIDATION = 32_767`.
- Support direct array in `validate_data(..., in: [...])` converting to `type: "list"` with quoted CSV `formula1`.
- Automatically infer `show_input_message: true` when `prompt:` or `prompt_title:` is provided, infer `show_error_message: true` when `error:` or `error_title:` is provided, and convert multiline prompts and errors (`\r\n`, `\n`) to OOXML `ST_Xstring` format (`_x000a_`).
- Support explicit `:default` / `:none` cell style identifier to emit `s="0"` and prevent unintended column style inheritance.
- Merge duplicate `sheet.column` definitions on the same column index and sort by column index in in-memory and streaming modes.
- Inherit `sheet_format[:default_col_width]` when a column defines `style:` without explicit `width:`.
- Support sheet visibility states (`state:` `:visible`, `:hidden`, `:very_hidden`, and `hidden: true`) across `StreamWriter#sheet`, `WorksheetProxy`, and `WorkbookBuilder#sheet`, propagate state into streaming workbook XML, reject invalid state values with `ArgumentError`, and enforce that at least one visible worksheet exists in `strict_excel_mode`.
- Add active sheet and workbook view configuration (`active_sheet`, `first_sheet`, `show_sheet_tabs`, `tab_ratio`, and `sheet(..., active: true)`) across `WorkbookBuilder`, `StreamWriter`, and `WorksheetProxy`, resolve active tab indices safely to avoid hidden sheet selection, emit `activeTab`, `firstSheet`, `showSheetTabs`, and `tabRatio` on `<workbookView>`, guarantee single `tabSelected="1"` emission on the active sheet, and expose `active_tab`, `active_sheet`, `first_sheet`, `show_sheet_tabs?`, and `tab_ratio` on `Elements::Workbook`.
- Enhance worksheet tab color configuration (`tab_color`) across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` with support for RGB hex strings, theme colors, tint, indexed colors, and auto color, emit full `<tabColor>` attributes in `WorksheetWriter` and `Ooxml::Writer`, and expose `tab_color` readers on `Elements::Worksheet` and `StreamSheet`.
- Support worksheet error check warning suppression rules (`<ignoredErrors>`) via `ignore_error` and `add_ignored_error` across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, emit schema-compliant `<ignoredErrors>` elements in `WorksheetWriter`, and expose `ignored_errors` readers on `Elements::Worksheet` and `StreamSheet` with full round-trip preservation.
- Support worksheet phonetic (furigana) guide properties (`<phoneticPr>`) via `phonetic_properties` (and alias `phonetic_pr`) across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, emit schema-compliant `<phoneticPr>` elements in `WorksheetWriter` and `WorkbookWriter` in Part 4 schema sequence position, and expose `phonetic_properties` query readers on `Elements::Worksheet` and `StreamSheet` with full round-trip preservation.
- Support workbook and worksheet code names (`code_name`), page setup fit-to-page linkage (`fit_to_page`, `fit_to_page?`), and worksheet calculation properties (`sheet_calc_pr`, `full_calc_on_load?`) across `WorkbookBuilder`, `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, emit schema-compliant `codeName` on `<workbookPr>` and `<sheetPr>`, emit `<pageSetUpPr fitToPage="1"/>` and `<sheetCalcPr fullCalcOnLoad="1"/>` in `WorksheetWriter` conforming to ECMA-376 Part 4 schema sequence, and expose query methods and properties on `Elements::Workbook`, `Elements::Worksheet`, and `StreamSheet` with full round-trip fidelity.
- Support worksheet background image (`<picture r:id="..."/>`) via `background_image` (with aliases `background_picture`, `picture`, `background_image=`, `background_picture=`, `picture=`) across `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, serialize relationship in `xl/worksheets/_rels/sheet{N}.xml.rels` and binary media in `xl/media/sheet{N}_bg.{ext}`, register image extension in `[Content_Types].xml`, emit schema-compliant `<picture>` element in `WorksheetWriter` conforming to ECMA-376 Part 4 schema sequence, and expose `background_image` query readers on `Elements::Worksheet` and `StreamSheet` with full round-trip fidelity.
- Add visual regression test example, Visual Gallery section, and WebAssembly Live Playground preset for worksheet views and advanced properties (`sheet_properties`).
- Support hierarchical differential format Hash (`dxf: { font: ..., fill: ... }`) in `conditional_format` and `add_conditional_format` across `StreamWriter`, `WorksheetProxy`, `WorksheetBuilder`, and `Ooxml::Writer` with automatic deduplication, emit `<b val="0"/>`, `<i val="0"/>`, `<strike val="0"/>` on boolean `false` font properties, and support automatic/theme/tint font and background colors and `pattern: "none"` pattern fills.
- Add `skip_rows` (and alias `skip_row`) to `StreamWriter`, `WorksheetProxy`, and `WorksheetBuilder` to advance the row cursor without writing empty `<row>` XML elements.
- Support configurable workbook-wide default font (`default_font: { name:, size: }`, `default_font(name, size, **opts)`, `set_default_font`) across `StreamWriter`, `WorksheetProxy`, `WorkbookBuilder`, `Xlsxrb.write`, and `Xlsxrb.build`, serializing name, size, family, charset, and scheme onto the base font (font 0).
- Support inline strings option (`shared_strings: false`) across `StreamWriter` and `Xlsxrb.write`, writing cells directly as `<c t="inlineStr"><is><t>...</t></is></c>` and omitting `xl/sharedStrings.xml` and package relationships to maintain flat $O(1)$ constant memory during streaming write.
- Support configurable ZIP deflate compression level (`compression_level: 0..9`) across `StreamWriter`, `Ooxml::ZipWriter`, `Ooxml::WorkbookWriter`, and `Xlsxrb.write`, allowing performance tuning between compression speed (e.g. level 4 for 20-40% faster writes) and archive size.

### Changed
- Unify differential format extraction across `WorkbookWriter` and `StreamWriter` into centralized `DslHelpers.extract_dxf_from_rule!`.
- Synchronize mutation testing metrics in `README.md`, `docs/MUTATION_TESTING.md`, and `docs/QUALITY_ASSURANCE.md` to reflect 51 pure functional subjects and 2,549 mutations at 100.00% kill rate.
- Optimize `bin/pre-push` and test suite execution time:
  - Scope mutation testing (`rake mutant:pure`) to subjects and test files matching changed files via `MUTANT_FILES` (running only affected pure logic subjects in ~2-20s rather than all 51 subjects), falling back to full coverage when `Rakefile` or `.mutant.yml` changes or when `--force` is passed, and suppressing parser future-frozen string warnings during mutation runs.
  - Eliminate duplicate `rake sig` invocation in `bin/pre-push` (already invoked as a prerequisite of `rake typecheck`) and run unit and contract tests in a single `rake test` process rather than two separate test processes.
  - Decouple serial RBS runtime type validation (`test:rbs`) from default local `bin/pre-push --all` (saving ~160s locally where static Steep type checking already ensures zero errors), while preserving automated runtime type validation on CI and allowing manual execution with `bin/pre-push --all --rbs` or `RUN_RBS=1`.
  - Condition SimpleCov coverage profiling on `ENV["COVERAGE"] == "1"` or `ENV["CI"] == "true"`, eliminating 5-10s of instrumentation overhead during local test runs.
  - Replace anonymous keyword argument forwarding (`**`) in `StreamWriter` and `WorksheetBuilder` with explicit `**opts` to prevent Steep 2.1.0 internal `RuntimeError` crashes during type synthesis.
  - Update `bin/setup` to install `.git/hooks/pre-push` as a wrapper forwarding directly to `bin/pre-push --all`, ensuring `git push` runs extended diff-scoped mutation verification while staying synchronized with `bin/pre-push`.

### Fixed
- Guard `Ooxml::Writer#generate_table_xml` against `TypeError` when `tbl[:style]` is provided as a string, and resolve table style attributes (`show_first_column`, `show_last_column`, `show_row_stripes`, `show_column_stripes`) directly from table definition root.
- Omit `<autoFilter>` tag from table XML when `auto_filter: false` is explicitly specified across `Ooxml::Writer` and `WorkbookWriter`.
- Enforce strict Excel worksheet name rules in `StreamWriter#sheet`, `WorkbookBuilder#sheet`, and `Elements::Worksheet.valid_name?` / `validate`, rejecting colons (`:`), leading/trailing apostrophes (`'`), and the case-insensitive reserved name `History`, updating validation error messages to reference Excel constraints.
- Guard against non-finite Float values (`Float::INFINITY`, `Float::NAN`, and non-finite cached formula values) in `StreamWriter#row` and `WorksheetWriter#write_row_values`, raising `ArgumentError` to prevent corrupt XML output.
- Support `xsd:boolean` representations (`"true"`, `"false"`, `"1"`, `"0"`) across `StylesParser`, `StylesListener`, and column definitions in `WorksheetParser`.
- Safely handle `nil` shared strings table index (`sst_index: nil`) across unstyled fast path and general cell paths in `WorksheetWriter#write_row_values`, writing string values as inline strings (`t="inlineStr"`) without error.
- Include `test/ooxml/**/*_test.rb` in Rake test tasks and update custom number format assertions to use `Elements::Styles#num_fmt`.
- Preserve formatted text runs (`<r><rPr>...</rPr><t>...</t></r>`) in inline strings (`<c t="inlineStr">`) for `Elements::RichText` cells across unstyled fast path and general row paths in `WorksheetWriter#write_row_values` when `shared_strings: false` is used.

## [0.1.16] - 2026-10-24

### Changed
- Encapsulate `Elements::Styles` with dedicated collection readers (`fonts`, `fills`, `borders`, `cell_xfs`, `num_fmts`, `dxfs`, `cell_style_xfs`, `cell_styles`, `table_styles`, `indexed_colors`, `mru_colors`, `colors`) and item lookups (`font`, `fill`, `border`, `cell_xf` / `xf`, `num_fmt`, `dxf`, `cell_style_xf`, `cell_style`), removing public Hash delegation (`[]`, `[]=`, `fetch`, `dig`, `keys`, `values`, `each`, `merge!`), and provide builder helpers (`add_font`, `add_fill`, `add_border`, `add_num_fmt`, `add_xf`, `add_dxf`, `add_cell_style_xf`, `add_cell_style`, `add_style`) that append entries, invalidate memoized caches, and return allocated IDs.
- Enforce binary encoding (`ASCII-8BIT`) on raw XML buffers in `StreamRow` and document byte vs character indexing consistency policy in `AGENTS.md` to prevent terminal row truncation or byte-offset drift when scanning documents with multibyte UTF-8 characters.
- Clarify streaming write memory profile and date/time timezone semantics in README and YARD documentation.
- Reading a file that is not a ZIP archive (including an empty file) raises `Xlsxrb::ZipError` instead of `ArgumentError`.
- Strict numeric cell values: an explicit numeric cell (`t="n"`) with an invalid number raises `Xlsxrb::ParseError`, while an empty `<v>` reads as `nil` instead of `0`, exponent notation such as `1E-3` reads as a Float, and non-numeric values in cells without `t` fall back to their string payload. Inline string and formula string cells without text read as `nil`, and `t="d"` cells read as their ISO 8601 text.

### Fixed
- Compound File Binary (CFB) header validation: Validate major version (3 or 4) and sector shifts in `Ooxml::Cfb::Reader#parse_header` and bounds-check sector reading against file size in `read_sector`, preventing `RangeError` integer overflows on corrupted or tampered inputs and raising `Xlsxrb::Error` gracefully.
- Unstyled empty cells and column ranges: Suppress style attribute (`s="..."`) emission on unstyled empty cells in `WorksheetWriter#write_row`, `write_cell`, and `write_row_values` to prevent unintended style inheritance explosion when rows or columns have styles, and preserve column `min`..`max` ranges on round-trip across `build_worksheet`, `Elements::Column`, and `WorksheetWriter#write_columns`.
- Data validation `showDropDown` and `imeMode` attributes: Serialize `showDropDown="1"` and `imeMode` on `<dataValidation>` elements in `WorksheetWriter#write_data_validations`, matching `Ooxml::Writer` and `DataValidationsListener`.
- Sheet views and selection serialization: Omit hardcoded `tabSelected="1"` in `WorksheetWriter#write_sheet_views` unless `tab_selected` is true, preventing multiple sheets from opening simultaneously selected in Excel, serialize full sheet view properties (`top_left_cell`, `zoom_scale_normal`, `view`, etc.), and preserve active cell selection across `WorksheetWriter`, `Elements::Worksheet`, `StreamSheet`, and `Xlsxrb.write`.
- Differential format (`<dxfs>`) parsing and boolean font properties: Parse differential formats (`<dxf>`) in `Ooxml::StylesParser` and serialize font, `numFmt`, fill, alignment, border, and protection in `WorkbookWriter`, parse and serialize ECMA-376 `CT_BooleanProperty` attributes (`<b val="0"/>`, `<i val="0"/>`) for font properties, and preserve `<color auto="1"/>` in `StylesParser`.
- Column and row style inheritance and serialization: Resolve named and inline column and row styles to integer style indices and propagate style inheritance to unstyled cells in `WorkbookBuilder#process_styles`, serialize row style attribute `s` on streaming `<row>` elements in `WorksheetWriter#write_row_values`, and guard `<col style>` emission against non-integer identifiers.
- Resolve boolean (`t="b"`) formula `cached_value` payloads (`"1"`, `"0"`, `"true"`, `"false"`) to `true` and `false` in `Ooxml::Reader::WorksheetListener#store_cell_value`, matching streaming and DOM cell parsers.
- Forward `:type`, `:ref`, and `:shared_index` (`:si`) formula attributes when building cells from a formula Hash (`{ formula: ..., type: :array, ref: ... }`) in `WorksheetBuilder#row`, `Xlsxrb.build_raw_cell`, and `Xlsxrb.build_raw_cell_from_value`.
- Unify `Ooxml::Reader#extract_zip_entry` and `Ooxml::Reader#entry_names` with `Ooxml::ZipReader` to support case-insensitive OPC part name resolution (per ECMA-376 Part 2 §9.1.1.2), ZIP64 archives, and data-descriptor streams uniformly across all reader APIs.
- Exclude non-worksheet relationship parts (`chartsheet`, `dialogsheet`, `macrosheet`, `xlMacrosheet`, `xlIntlMacrosheet`) and `xl/chartsheets/`, `xl/dialogsheets/`, `xl/macrosheets/` targets from standard worksheet enumeration in `Xlsxrb.read`, `Xlsxrb.modify`, and `Ooxml::Reader#discover_sheets`.
- Excel date boundary guards: Raise `ArgumentError` when writing dates before the date system epoch (`1899-12-31` / `1900-01-01` in 1900 date system or `1904-01-01` in 1904 date system) or after `9999-12-31` in `date_to_serial`, and raise `ArgumentError` when reading negative serials or serials exceeding `9999-12-31` (`serial >= 2_958_466` in 1900 system or `serial >= 2_957_004` in 1904 system) in `serial_to_date`, adhering to ECMA-376 date constraints and preventing corrupted Julian day calculations or negative date representation errors.
- Formula cached value falsy resolution and boolean typing: Fix falsy checks on formula `cached_value` across `Xlsxrb`, `build_raw_cell_from_value`, and `WorksheetWriter` so that boolean `false` and numeric `0` are not skipped or converted to `nil`, and serialize boolean cached values with cell type `t="b"` and `<v>1</v>` / `<v>0</v>` per ECMA-376 Part 4 §18.3.1.4.
- Boolean cell parsing tolerance: Support text payloads `"true"` and `"false"` (case-insensitive) in addition to numeric `"1"` and `"0"` for cells with boolean type attribute `t="b"`, conforming to ECMA-376 `xs:boolean` specification and preventing boolean `true` cells from evaluating to `false`.
- Template modification in `Xlsxrb.modify`: Support case-insensitive and slash-normalized archive entry replacement per ECMA-376 Part 2 §7.3, preventing shadowed or duplicate entries when archives use mixed casing (such as `xl/SharedStrings.xml`) or Windows backslashes (such as `xl\worksheets\sheet1.xml`). Preserve XML namespace prefixes (such as `<x:worksheet>`, `<x:sst>`, `<x:si>`) during sheetData replacement and shared strings appending, prevent duplicate `xmlns:r` attribute definitions in namespaced and ISO/IEC 29500 Strict documents, and correctly anchor `<hyperlinks>` and `<legacyDrawing>` after preceding root container elements to avoid illegal nesting inside `<customSheetViews>` or extension blocks.
- Resolve Open Packaging Conventions (OPC) workbook part path dynamically via root relationships (`_rels/.rels`), enabling reading and modifying workbooks with non-standard part names or root locations (such as `xl/workbook2.xml` or `workbook.xml`), and fall back to standard sheet and part locations when workbook relationships are missing or corrupted.
- Handle multibyte UTF-8 characters encoded as ASCII-8BIT (binary) strings during XML escaping and buffer concatenation in worksheet writer, workbook writer, and streaming sheet, preventing `Encoding::CompatibilityError` when cell values or formulas contain multibyte text.
- ECMA-376 schema conformance during workbook serialization: Format hash-based colors (`{ theme: ... }`, `{ indexed: ... }`, `{ auto: ... }`) into schema-compliant `<color>` elements in stylesheets and workbooks instead of emitting unparseable Hash string representations (`rgb="{...}"`), strictly validate 6/8-character hex colors against `ST_UnsignedIntHex`, reorder DrawingML line and character properties (`a:ln`, `a:rPr`) to satisfy ECMA-376 `xs:sequence` child constraints, preserve leading and trailing whitespace using `xml:space="preserve"` on `<si><t>` elements in `sharedStrings.xml`, drop data validations missing required `sqref` references, and omit negative `firstPageNumber` values in worksheet `page_setup`.
- String fallback when cell type attribute `t` is omitted: Parse numeric strings (integers, floats, exponents) when cell type attribute is omitted, while treating non-numeric text payloads in `<v>` as raw strings rather than raising `Xlsxrb::ParseError`, allowing parsing of cells generated by Open XML SDK and other spreadsheet producers without crashing.
- XML namespace prefix tolerance in shared strings parser (`Ooxml::SharedStringsParser`): Locate `<sst>`, `<si>`, `<t>`, and `<rPh>` tags dynamically regardless of XML namespace prefixes (such as `<x:sst>`, `<x:si>`, `<x:t>`), preventing dropped shared strings in spreadsheets produced by Microsoft Open XML SDK, ClosedXML, and other XML-prefixed generators.
- Consolidate cell value interpretation across streaming fast-scan direct, row values, events, and SAX listener parsers into a centralized `Ooxml::WorksheetParser.resolve_cell_value` helper, eliminating redundant logic and preventing parser divergence.
- Preserve worksheet and workbook attributes during in-memory writes (`Xlsxrb.write(workbook)`): accurately map and serialize `freeze_pane`, `split_pane`, `sheet_view`, `page_setup`, `page_margins`, `header_footer`, `print_options`, `sheet_protection`, `tables`, `merged_cells`, `auto_filter`, `row_breaks`, `col_breaks`, and embedded images from `Elements::Worksheet`, and propagate workbook properties (`core_properties`, `app_properties`, `custom_properties`, `workbook_protection`).
- Support arbitrary XML attribute ordering on cells (such as `t` before `s` or `r` after `s`) in streaming fast-scan parsers, ensuring cell types and styles are never dropped regardless of producer attribute sequence.
- Empty self-closing string cells without `<v>` in fast scan parsers evaluate to `nil` instead of empty string `""`.
- In-memory SAX listeners (`WorksheetListener` and `SharedStringsListener`) skip phonetic readings (`<rPh>`) and treat empty self-closing string cells (`<c ... t="s"/>`) as nil, matching streaming reader behavior.
- Rich text parsing in shared strings parser (`Ooxml::SharedStringsParser`): Handle self-closing `<t/>` tags when the first run is empty without dropping subsequent runs, while bounding phonetic tag scans to maintain linear throughput.
- `extract_inline_text` and `extract_multi_t` strictly distinguish `<t>` elements from other XML tags containing 't' (such as `<rPr>`, `<vertAlign>`, or `<table>`), preventing corrupted inline string parsing in rich text runs.
- Formulas are serialized strictly when explicitly defined via `Xlsxrb.formula(...)` or `{ formula: ... }`. Plain string cells starting with `=` are always serialized as text cells, preventing formula injection vulnerabilities and unintended 0-value coercions on reload.
- Streaming reads no longer drop the cell that follows a self-closing empty cell (`<c r="A1" s="1"/>`) in `StreamRow#values` and `StreamRow#cells`.
- Streaming reads place a row-leading cell with a multi-letter column reference (`AB1`) in its own column instead of column A.
- Phonetic readings (`<rPh>`) are excluded from shared strings and inline strings, so cells with furigana no longer concatenate phonetic readings to display text.

### Added
- Extract and preserve `activeTab` (`<workbookView>`) and `fullCalcOnLoad` (`<calcPr>`) in `Ooxml::WorkbookParser`, `Xlsxrb.read`, and `Xlsxrb.modify`, and expose `workbook_properties`, `active_tab`, and `full_calc_on_load?` (`full_calc_on_load`) readers on `Elements::Workbook`.
- Support self-closing row tags (`<row r="N" .../>`): Extract row indices and custom row attributes (`ht`, `hidden`, `customHeight`, `outlineLevel`, `collapsed`, `s`) and emit `:row_start` / `:row_end` events and empty `StreamRow` instances across direct and streaming worksheet parsers.
- ISO/IEC 29500-1 ST_Xstring control character decoding: Support decoding `_xHHHH_` hexadecimal escape sequences (such as `_x000D_` carriage returns and `_x005F_` literal underscores) in `XmlBuilder.unescape`, `SharedStringsParser`, and `SharedStringsListener`, correctly normalizing newlines and preserving control character fidelity in cell strings.
- Record Architecture Decision Record [ADR 0001: Excel Date System Boundary Handling](docs/adr/0001-excel-date-system-boundary-handling.md) documenting date conversion semantics, ecosystem survey, and design rationale across 1900 and 1904 date systems.
- Date validation and boundary checks: Add `Xlsxrb::Ooxml::Utils.valid_excel_date?` (and delegates on `Xlsxrb::Utils` and `Xlsxrb`) to check if a numeric serial value falls within the supported Excel date boundaries [0, 2,958,465] (1900 system) or [0, 2,957,003] (1904 system).
- `ZipReader#open_entry_io` and `ZipReader#each_entry_chunk` support an `encoding:` parameter and automatically resolve `.xml` and `.rels` entries to UTF-8 encoding, preventing `Encoding::CompatibilityError` in downstream text and regular expression operations.
- `Xlsxrb::Ooxml::Utils.booleanize` and `Xlsxrb::Ooxml::Utils.camelize` helper methods for OOXML attribute serialization and name casing conversion.
- Round-trip contract tests in `test/contract/round_trip_fidelity_test.rb` verifying that `read -> write -> read` preserves worksheet and workbook attributes across streaming and loaded in-memory representations.
- Alias `conditional_formatting` to `conditional_formats` on `StreamSheet` for complete API symmetry with `Elements::Worksheet`.
- Property-based tests for cell XML parsing in `test/pbt/cell_xml_parser_pbt_test.rb`, validating 100% equivalence between streaming fast-scan and SAX listener parsers across random attribute permutations, namespaces, rich text, and phonetic elements.
- Dedicated exception `Xlsxrb::ZipBombError < Xlsxrb::ZipError` for archive expansion limit violations.
- Workbook view and calculation properties in in-memory writes: `Xlsxrb.write(workbook)` emits `<bookViews><workbookView activeTab="n"/></bookViews>` and `<calcPr fullCalcOnLoad="1"/>` when the workbook properties carry `active_tab:` and `full_calc_on_load: true`.
- Autonomous QA agent squad skill definitions and documentation: Added skill playbooks in `.agents/skills/` (`throughput-performance`, `security-resilience`, `api-contract`, `concurrency-safety`, `visual-fidelity`), and updated `docs/QA_AGENTS.md`, `docs/QUALITY_ASSURANCE.md`, `.agents/rules/qa_agents.md`, and `AGENTS.md`.
- Visual fidelity and Wasm compatibility verification agent (`visual-fidelity`): Formulated autonomous skill playbook covering headless LibreOffice Calc visual regression testing (`rake test:visual`) and WebAssembly playground asset verification (`rake doc:check`).
- API symmetry and contract verification agent (`api-contract`): Added reflective method signature and parameter arity parity tests between `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` in `test/contract/api_symmetry_test.rb`, added `very_hidden?` predicate on both `Elements::Worksheet` and `StreamSheet`, and enforced lossless round-trip serialization contracts between streaming and in-memory workflows.
- Security and resilience verification agent (`security-resilience`): Added regular expression ReDoS audit tests in `test/security_resilience_test.rb` verifying linear processing under 10k+ character adversarial inputs, supported `max_uncompressed_size` option on `Xlsxrb.read` block form to protect streaming decompression from ZIP bombs, and added `escape_formulas: true` option on `Xlsxrb.write`, `StreamWriter`, and `WorksheetWriter` to prevent formula and CSV injection vulnerabilities.
- Throughput and complexity verification agent (`throughput-performance`): Added deterministic allocation bounds testing (< 25 objects/row write, < 15 objects/row read for 10k rows) and scaling ratio tests in `test/performance/throughput_test.rb`, hot-path microbenchmarks with `benchmark-ips` (`benchmark/micro_ips.rb`), and corresponding `benchmark:ips` rake task.
- Concurrency and Ractor safety verification agent (`concurrency-safety`): Added automated reflection scan testing that 100% of constants in the `Xlsxrb` namespace are `Ractor.shareable?`, deeply froze `EPOCH_1900`, `EPOCH_1904`, `INTEGER_STRINGS`, and `LOCAL_HEADER_SIG`, added Ractor-safe accessors for REXML security settings in `XmlParser`, and added full parallel write/read/modify pipeline verification across concurrent Ractors in `test/concurrent_test.rb`.

## [0.1.15] - 2026-10-17

### Added
- Block form for table and data validation configuration: Support optional configuration block in `table` and `validate_data` methods on `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` using `TableBuilder` and `DataValidationBuilder` contexts.
- Protected range and cell watch configuration DSL and properties (`protected_range` and `cell_watch`): Add `protected_range` (and alias `add_protected_range`) and `cell_watch` (and alias `add_cell_watch`) methods to `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy`, normalize passwords into ECMA-376 compliant hashes or modern SHA-512 protection attributes, serialize `<protectedRanges>` and `<cellWatches>` according to OOXML schema sequence rules in `WorksheetWriter`, and expose `protected_ranges` and `cell_watches` readers on `Elements::Worksheet` and `StreamSheet`.
- Range coordinate conversion utilities (`Xlsxrb::Utils.parse_range` and `range_to_ref`): Parse cell range references (`"A1:C10"`, `"A1"`) to 0-based `[[top_row, left_col], [bottom_row, right_col]]` coordinate pairs, and format bounding coordinate tuples into A1-style range references (`"A1:C10"`) with `Xlsxrb.parse_range` and `Xlsxrb.range_to_ref` facade methods.
- Sheet format configuration DSL and properties (`sheet_format`): Add `sheet_format` method to `WorksheetBuilder`, `StreamWriter`, and `WorksheetProxy` supporting `default_row_height:`, `default_col_width:`, and extended properties (`outline_level_row`, `outline_level_col`, `zero_height`, `custom_height`, `thick_top`, `thick_bottom`), serialize `<sheetFormatPr>` schema-compliantly in `WorksheetWriter` and `WorkbookWriter`, and expose `sheet_format` reader on `Elements::Worksheet` and `StreamSheet`.
- Numeric and hexadecimal XML entity decoding in `XmlBuilder.unescape`: Support decoding decimal (`&#...;`) and hexadecimal (`&#x...;`) XML numeric character references with UTF-8 encoding and `RangeError` fallback.
- ZIP64 extra field support in `ZipReader`: Support reading ZIP64 extra fields (`0x0001`) in both central directory headers and local file headers for 64-bit uncompressed sizes, compressed sizes, and local header offsets.
- Expose `merged_cells` and `auto_filter` on `Elements::Worksheet`: Expose parsed sheet merge ranges and auto filter definitions directly on in-memory worksheets when loaded via `Xlsxrb.read` and `StreamSheet#load`.
- Expose document properties and protection settings on `Elements::Workbook`: Expose `core_properties`, `app_properties`, `custom_properties`, and `workbook_protection` readers on in-memory workbooks, populating them during `Xlsxrb.read` from `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml`, and `<workbookProtection>`.
- Support updating cell hyperlinks and comments in `Xlsxrb.modify`: Synchronize modified cell hyperlinks and comments during template modification, updating or creating `_rels/sheetN.xml.rels`, `commentsN.xml`, `vmlDrawingN.vml`, and schema-compliant `<hyperlinks>` and `<legacyDrawing>` tags in worksheet XML.
- Expose sheet structure, print, and view metadata on `Elements::Worksheet` and `StreamSheet`: Expose `tables`, `shapes`, `freeze_pane`, `page_setup`, `page_margins`, `header_footer`, `print_options`, `sheet_protection`, `sheet_view`, `sheet_properties`, `row_breaks`, and `col_breaks` on both streaming and in-memory worksheets.
- WebAssembly Live Playground Monaco Editor integration: Replaced `<textarea>` with Monaco Editor in `docs/preview.html` featuring `vs-dark` theme, Ruby syntax highlighting, 2-space auto-indentation, auto-closing brackets and quotes, minimap, line numbers, and xlsxrb DSL snippet completion candidates (`Xlsxrb.write`, `wb.sheet`, `wb.style`, `s.row`, `s.column`, `s.chart`, `s.table`, `s.merge`, etc.). Implemented Blob URL worker initialization for Cross-Origin Isolation (COOP/COEP) compatibility and isolated AMD loader state during UMD binary bootstrap.
- Automated documentation and playground verification (`rake doc:check`): Validate generated documentation assets (file existence, minimum sizes, and magic bytes), verify HTTP server responses (status 200/206, MIME types, COOP/COEP headers, and absence of invalid compression headers) via byte-range requests, and automatically parse and execute all Live Playground presets to guarantee error-free generation prior to preview and deployment.
- WebAssembly Live Playground: Revamped `docs/preview.html` with current `Xlsxrb.write` DSL, added visually striking Executive Dashboard first view featuring an embedded Ruby gem logo image, KPI summary cards, zebra-striped table, currency formats, `=SUM` formulas, and a clustered bar chart in ~55 lines, added interactive preset selector with 11 spreadsheet scenarios (executive dashboard, sales reports, bar/line/pie charts, tables, borders, number formats, conditional formatting, hyperlinks, and Japanese text), direct URL deeplinking (`?example=...`), and seamless link integration throughout `VisualGallery.md` and `README.md`.
- New visual examples and VRT baselines for missing formatting and print features: `hyperlinks.rb` (external and internal cell hyperlinks), `column_styles.rb` (column-level default number formats and styling), `page_repeat_rows.rb` (repeating header rows for print setup), and `sheet_protection.rb` (password protection with granular permissions).

### Changed
- Standardized top navigation header across all documentation pages and Live Playground: Unified navbar with `Documentation`, `Visual Gallery`, `Live Playground`, and `GitHub` navigation links, active page highlighting, search input, and SVG theme toggle.
- Reorganized documentation sidebar navigation: Promoted `Visual Gallery` and `Live Playground` to prominent action items at the top of the sidebar, renamed the collapsible `docs` folder to an open `Guides` section, and formatted screaming uppercase filenames into clean Title Case labels (`Architecture`, `Development Guide`, etc.) to eliminate visual noise.
- Enhanced search index discoverability: Pre-indexed all documentation guides, pages, and the Live Playground in `search_data.js` so search queries for "Visual Gallery", "Playground", and "Architecture" yield instant results.
- Universal rendering without font tofu: Replaced theme toggle emojis with SVG sun and moon icons, and replaced README comparison table emoji checkmarks with accessible text badges (`Yes` / `-`) to eliminate tofu glyphs (`□`) across Linux and headless environments.
- Visual Gallery layout polish: Constrained gallery overview thumbnails to 160px width with neat borders, styled `Live Playground` actions as interactive buttons, and constrained full baseline preview images to a maximum height of 480px to avoid unbounded scrolling whitespace.
- Documentation callout banners: Transformed GitHub-style markdown alert blockquotes (`[!TIP]`, `[!NOTE]`, `[!WARNING]`) into styled callout banners with SVGs.
- Cleaned up duplicated inline doc tags: Stripped repeated `rbs_inline: enabled` and `@api public` tags from generated module documentation.
- Align WebAssembly Live Playground design with documentation: Unify `docs/preview.html` visual styling with the RDoc Aliki documentation design system, adding light and dark theme support with local storage persistence, synchronized Monaco Editor theme switching (`vs` and `vs-dark`), top navigation bar matching documentation layout with links to documentation and visual gallery, solid brand coral red Run button matching documentation accent tokens, and dynamic status dot indicator (`Initializing...`, `Ready`, `Rendering...`, `Error`).
- Cleaned up visual examples and gallery generator: Stripped internal read validation debug dumps across all visual examples to present concise DSL code snippets, refreshed formula and merge syntax to use modern public APIs (`Xlsxrb.formula`, `A1:C1` ranges), and reorganized `VisualGallery.md` into 9 structured capability categories with dedicated preview tables and live playground links.
- Removed non-visual and redundant visual examples: Removed `invoice_statement.rb` and `streaming_append.rb` and their baseline directories from the visual regression test suite and gallery.
- Use `nil` for unstyled cells in `embedded_images.rb` example instead of undefined `"left"`.

### Fixed
- XML namespace prefix tolerance in template modification (`replace_sheet_data`): Locate `<sheetData>` and matching closing tags dynamically regardless of XML namespace prefixes (such as `<x:sheetData>` or `<d:sheetData>`) or self-closing syntax during `Xlsxrb.modify`.
- Resilience against duplicate XML declarations and malformed `.rels` files in `RelationshipsParser`: Sanitize duplicate or multiple `<?xml ...?>` declarations and provide regex fallback scanning if SAX parsing fails on malformed relationship XML.
- Fix false positive date detection for color tags like `[Red]` in `Ooxml::Reader#date_pattern?`: Unify custom number format date evaluation with `NumberFormatter.date_format?`, ignoring bracketed formatting codes (such as `[Red]` or `[Green]`) so negative numbers with colored styles are not misidentified as dates.
- Fix layout overlap between Playground Run & Preview button and theme toggle: Override absolute positioning of `.theme-toggle` from RDoc Aliki styles with static inline-flex flow in `docs/preview.html` to eliminate button collision and preserve clean spacing across light and dark modes.
- Prevent WebAssembly out-of-memory error during ZetaOffice instantiation on Firefox: Provide custom `Module.instantiateWasm` handler in `docs/preview.html` with streaming-first compilation (`WebAssembly.instantiateStreaming` / `WebAssembly.compileStreaming`) to bypass Emscripten's default `instantiateArrayBuffer` fallback that allocates an extra 160MB ArrayBuffer on memory pressure, and sequence Ruby Wasm and LibreOffice Wasm engine initialization to eliminate concurrent memory peaks.
- Resolve WebAssembly Live Playground (`docs/preview.html`) console errors and warnings across Chromium and Firefox:
  - Compile `ruby.wasm` via streaming (`WebAssembly.compileStreaming`) with automatic fallback to avoid large intermediate ArrayBuffer heap allocations, eliminating `InternalError: out of memory` on memory-constrained browser environments.
  - Bypass service worker registration and reloading in `coi-serviceworker.js` when native Cross-Origin Isolation headers (`COOP`/`COEP`) are already present, preventing in-flight network request abortions (`NS_BINDING_ABORTED`).
  - Filter benign C/C++ Emscripten/Qt runtime stderr diagnostics in `Module.printErr` and route diagnostic output to `console.log` rather than `console.warn`.
  - Guard UNO document model close/dispose handling in `office_thread.js` to cleanly close previous models without emitting `console.warn` on preset switching or re-runs.
- Prevent unhandled clipboard permission errors in Live Playground on Firefox: Intercept `navigator.permissions.query` in `docs/preview.html` and worker thread bootstrap to resolve dummy `PermissionStatus` objects for `clipboard-read` and `clipboard-write`, preventing uncaught `TypeError` exceptions from Embind in browsers that do not support clipboard descriptors in `PermissionDescriptor`.
- Fix missing closing angle bracket on styled formula cells in `WorksheetWriter`: Add missing `>` before `<f>` element when emitting styled formula cells in `WorksheetWriter#write_row_values`, resolving XML syntax errors (`<c r="..." s="..."<f>`) that caused LibreOffice and spreadsheet engines to reject sheet formatting.
- Prevent serialization of unresolved non-integer style names as `s` attribute in `WorksheetWriter`: Guard style resolution in `WorksheetWriter` to only emit valid integer style indices, ignoring unresolved strings or symbols to maintain ECMA-376 XML Schema conformance (`xs:unsignedInt`).
- Prevent invalid `Content-Encoding: br` header in local `doc:preview` server: Remove erroneous Brotli encoding header injected on uncompressed Emscripten `.data` packages (such as `soffice.data`), and supply native Cross-Origin Isolation headers (`COOP`, `COEP`, `CORP`) to resolve browser decompression failures (`NS_ERROR_INVALID_CONTENT_ENCODING`).

## [0.1.14] - 2026-10-10

### Added
- Non-self-closing `<sheet>` and `<xf>` tag compatibility: Verify and support spreadsheet XML where `<sheet>...</sheet>` in `workbook.xml` or `<xf>...</xf>` in `styles.xml` are serialized with explicit closing tags rather than self-closing syntax.
- Namespace prefix resilience for relationship and embed attributes: Match relationship and embed attributes by local name or any namespace prefix (`r:id`, `d3p1:id`, `rel:id`, `id`, `r:embed`, `embed`) across `WorkbookListener`, `HyperlinksListener`, `DrawingListener`, and `WorksheetParser`, ensuring reliable sheet, hyperlink, and media resolution for non-standard XML generators.
- UTF-16LE, UTF-16BE, and UTF-8 BOM encoding normalization in `ZipReader`: Automatically detect and transcode Byte Order Mark (BOM) encoded XML streams (`\xFF\xFE`, `\xFE\xFF`, and `\xEF\xBB\xBF`) to UTF-8 in `ZipReader#read_entry` and `EntryStreamIO`, enabling seamless parsing of workbooks generated on Windows Phone and legacy platforms.
- XML 1.0 line ending normalization (`\r\n` and `\r` to `\n`): Normalize line breaks in shared strings, inline strings, and cell text values per W3C XML 1.0 §2.11 across `SharedStringsParser` and `WorksheetParser`.
- Whitespace and newline tolerance in cell regex scanning: Allow whitespace and newline sequences between `<c>` child tags (`<f>`, `<v>`, `<is>`) and strip whitespace on extracted values in `WorksheetParser`, ensuring reliable parsing of pretty-printed or formatted worksheet XML.
- Sequential row index fallback for omitted `r` attributes: Track previous row index and fall back to sequential 1-based indexing when `<row>` elements omit the `r` attribute per ECMA-376 Part 1 §18.3.1.73 across direct string scanning, chunked stream scanning, and event scanning in `WorksheetParser`.
- Case-insensitive part name lookup in `ZipReader`: Fall back to case-insensitive lookup for package part entries per ECMA-376 Part 2 §8.1.1, allowing workbooks with case-varying part names (such as lower-case `xl/sharedstrings.xml` or `xl/styles.xml`) to resolve transparently.
- Windows backslash path separator normalization in `ZipReader`: Normalize Windows backslash separators (`\`) in ZIP central directory and local headers to standard forward slashes (`/`), enabling transparent extraction and catalog lookup for XLSX archives generated on Windows.
- First-class sparkline groups on worksheet and `add_sparkline` helper: Expose `sparkline_groups` reader and keyword argument on `Elements::Worksheet`, propagate `ws.sparkline_groups` during `Xlsxrb.write` DOM serialization, and add `add_sparkline(location:, range:, type:, **opts)` helper to `WorksheetBuilder`, `StreamWriter`, and `StreamWriter::SheetContext` supporting `:win_loss` mapping to `"stacked"`.
- High-level print area and print titles DSL: Add `print_area`, `print_titles`, `repeat_rows`, and `repeat_columns` to `WorksheetBuilder`, `StreamWriter::SheetContext`, and `StreamWriter`, expose `print_area` and `print_titles` on `Elements::Worksheet`, and auto-generate `_xlnm.Print_Area` and `_xlnm.Print_Titles` defined names in `WorkbookBuilder#build` and `Xlsxrb.write` DOM export.
- ECMA-376 indexed color palette and `Xlsxrb::Colors`: Add `Xlsxrb::Colors` providing `INDEXED_COLORS` (indices 0..63 per ECMA-376 Part 4 §18.8.27), `palette_color`, and `to_hex` utilities, expose `Xlsxrb.palette_color` and `Xlsxrb.to_hex_color` facade methods, and resolve integer palette indices 0..63 automatically in `StyleBuilder.resolve_color`.
- Flexible rich text run input in `Xlsxrb.rich_text`: Support `[text, font_hash]` pair arrays, alternating varargs of plain strings and format hashes or objects responding to `to_font_hash`, and mixed `Elements::RichTextRun` instances.
- First-class `Elements::Hyperlink`, `Xlsxrb.hyperlink`, and `Cell#tooltip`: Add `Elements::Hyperlink` data class and `Xlsxrb.hyperlink` factory helper with `tooltip:` support, expose `Cell#tooltip` and `:tooltip`, `:url` symbol key access on `Elements::Cell`, and support `Elements::Hyperlink` instances in `WorksheetBuilder#hyperlink`, `StreamWriter#hyperlink`, and `Ooxml::Writer#hyperlink`.
- Array and dynamic array formula serialization and `Xlsxrb.array_formula`: Add `Xlsxrb.array_formula(expression, ref, cached_value:)` factory method, extend `Xlsxrb.formula` with optional `type:` and `ref:` keywords, and serialize `<f t="array" ref="...">` and shared formula attributes across all `WorksheetWriter` serialization paths.
- First-class `style_index:` keyword on `Elements::Column`: Add `style_index:` keyword argument directly to `Elements::Column.new`, storing it in `Data.define` attributes and populating it from `rc[:style_index]` during worksheet parsing and DSL column definitions.
- In-memory streaming write and row append chaining: Support `Xlsxrb.write` with omitted target in block mode to return an in-memory binary String, support self-returning `<<` row appending for chaining across `StreamWriter`, `SheetContext`, and `WorksheetBuilder`, and add `add_row` aliases.
- Unified color and border style normalizers in `StyleBuilder`: Add `StyleBuilder.normalize_color` (supporting 140+ CSS named colors, `#RGB`, `#RRGGBB`, `#AARRGGBB` hex strings, numeric RGB integers, and optional alpha channel) and `StyleBuilder.normalize_border_style` (supporting border style symbols, aliases, strings, and 0..13 numeric border indices).
- Public date and time serial conversion utilities: Expose `date_to_serial`, `serial_to_date`, `datetime_to_serial`, and `serial_to_time` (aliased `serial_to_datetime`) on `Xlsxrb` and `Xlsxrb::Utils` for converting Ruby `Date` and `Time` instances to and from Excel 1900 and 1904 serial numbers.
- Cascade style inheritance and column styles: Support cascade style inheritance (Cell Style > Row Style > Column Style) in `WorksheetWriter`, and add `style:` parameter to `StreamWriter#column` and `WorksheetBuilder#column`.
- Embedded drawing and image extraction: Add `Elements::Image` representing anchored embedded images, expose `images`, `images_at`, and `with_images` across `StreamSheet` and `Elements::Worksheet`, and expose `images` and `images_at` across `Elements::Workbook`.
- Configurable sparse vs padded row and cell iteration: Add `pad_empty_rows` and `pad_empty_cells` options to `Xlsxrb.read`, `Xlsxrb.build_worksheet`, `StreamSheet`, `WorksheetParser.each_row`, and `Elements::Worksheet` to pad skipped rows or missing cells up to the maximum column, and add `StreamRow#to_h` and `Elements::Row#cells_hash` for coordinate-keyed row maps.
- Row XML metadata accessors and serialization: Expose row XML metadata (`collapsed`, `collapsed?`, `hidden?`, `custom_height?`, `attributes`) on `StreamRow` and `Elements::Row`, support `collapsed:` in `StreamWriter#row` and `WorksheetBuilder#row`, and serialize `collapsed="1"` in `WorksheetWriter`.
- Format classification helpers: Add `Elements::Styles#format_code_for` and `#classification_for` aliases, allow `NumberFormatter` format predicates (`date_format?`, `time_format?`, etc.) to resolve built-in formats by `num_fmt_id`, and expose `format_type`, `date_format?`, and `time_format?` on `Elements::Cell`.
- Date1904 accessor: Expose `Ooxml::Reader#date1904?` and add `date1904:` keyword argument to `Elements::Workbook#initialize`.
- Coordinate utilities module: Add `Xlsxrb::Utils` exposing public coordinate helpers (`ref_to_row_col`, `row_col_to_ref`, `col_name_to_index`, `col_index_to_name`, `split_coordinate`).
- Configurable trailing empty row trimming: Add `trim_empty_rows` option across `Xlsxrb.read`, `Xlsxrb.build_worksheet`, `WorksheetParser.each_row`, `StreamSheet#initialize`, `StreamSheet#each_row`, `StreamSheet#each_row_values`, `StreamSheet#each`, `Elements::Worksheet#initialize`, `Elements::Worksheet#each_row`, and `Elements::Worksheet#each_row_values` to omit trailing empty rows, and add `empty?` predicate to `StreamRow` and `Elements::Row`.
- Direct typed values streaming iterator: Add `StreamSheet#each_row_values` and `Elements::Worksheet#each_row_values` to stream row values directly as Arrays without allocating `Elements::Cell` objects, and add `type_cast: true` support across `StreamSheet#each_row_values`, `Elements::Worksheet#each_row_values`, `StreamRow#values`, and `Elements::Row#values` to coerce date/time serial numbers into `Date` and `Time` instances.
- Precomputed number format classifications in Styles: Add `Xlsxrb::Elements::Styles` (subclass of `Hash`) with precomputed $O(1)$ format classification lookups (`number_format`, `format_type`, `date_format?`, `datetime_format?`, `time_format?`, `date_only_format?`), update `StylesParser.parse` to return `Elements::Styles`, and add classification helper methods to `Xlsxrb::NumberFormatter`.
- Worksheet dimension and coordinate bounds: Expose `dimension`, `first_row`, `first_column` (aliased `first_col`), `last_row`, and `last_column` (aliased `last_col`) on `StreamSheet` and `Elements::Worksheet` (and `Elements::CoordinateAccess`), lazily extracting `<dimension ref="..."/>` in $O(1)$ memory without scanning the full sheet.
- Date1904 propagation to streaming sheets and rows: Propagate the workbook's 1904 date system flag to `StreamSheet` and `StreamRow`, expose `date1904?` on `StreamSheet`, `StreamRow`, `Elements::Worksheet`, and `Elements::Cell`, and default `Cell#to_date` and `Cell#to_time` to the cell's `date1904` setting.
- `Elements::RichTextRun`: Add `RichTextRun` data class representing formatted text runs within `RichText`, supporting hash-like bracket access and seamless integration with `Xlsxrb.rich_text`.
- First-class conditional formatting and data validation on `Elements::Worksheet`: Expose `conditional_formatting` (and alias `conditional_formats`) and `data_validations` as first-class fields on `Elements::Worksheet`, populate them during round-trip parsing in `Xlsxrb.build_worksheet`, and serialize them in in-memory workbook serialization.
- Date1904 calendar system arithmetic: Support the 1904 date system (1904-01-01 base epoch) across serial date conversions, `Cell#to_date`, `Cell#to_time`, `Elements::Workbook#date1904?`, `Ooxml::Reader`, `Ooxml::WorksheetWriter`, `Ooxml::WorkbookWriter`, and `StreamWriter`.
- Row and column styles persistence: Expose `style_index` on `Row`, `Column`, and `StreamRow`, and persist row and column styles (`<row s="..." customFormat="1">` and `<col style="..."/>`) during serialization and round-trip parsing.
- Styles metadata persistence: Support border colors, cell alignment, and protection attributes in stylesheet generation and round-trip parsing.
- Dual format support for `num_fmts`: Support both `Hash` (`{ 164 => "format" }`) and `Array` (`[{ num_fmt_id: 164, format_code: "format" }]`) structures for custom number formats in `WorkbookWriter`.
- Initial shared strings table seeding: Support pre-seeding `sst` in `Xlsxrb.write` when `workbook.shared_strings` is provided.
- Sheet visibility state: Support worksheet visibility state (`state`, `hidden?`, `visible?`) across `StreamSheet`, `Elements::Worksheet`, `WorkbookParser`, `WorkbookBuilder`, and `WorkbookWriter`, preserving `:visible`, `:hidden`, and `:very_hidden` visibility in streaming read, in-memory representations, and XML serialization.
- Workbook defined names: Parse `<definedNames>` in `WorkbookParser` and expose `defined_names` array and `defined_name(name, sheet: nil)` lookup on `Elements::Workbook`, supporting global and sheet-scoped named ranges and formula constants.
- Cell formula query and expression methods: Expose `formula?` predicate and `formula_expression` string accessor on `Elements::Cell`, and normalize empty formula expressions to `nil` in streaming and DOM cell parsing.
- Worksheet and cell hyperlinks: Parse and expose hyperlink metadata across `StreamSheet#hyperlinks`, `Elements::Worksheet#hyperlinks`, and `Elements::Cell#hyperlink`, `#link?`, and `#url`, supporting external targets and internal sheet references in streaming read, DOM models, and builders.
- Worksheet and cell comments: Parse and expose comment and note metadata across `StreamSheet#comments`, `StreamSheet#comment`, `Elements::Worksheet#comments`, and `Elements::Cell#comment`, `#comment?`, and `#comment_text`, supporting comment annotations and coordinate lookups in streaming read, DOM models, and builders.
- Number formatting engine and extended cell metadata: Add `Xlsxrb::NumberFormatter` supporting standard and custom OpenXML format codes (currencies, percentages, date/time, fixed-point padding, accounting patterns, scientific notation, and text), expose `formatted_value`, `raw_value`, and `format_code` on `Elements::Cell`, add `formatted_value` coordinate lookups on `Elements::Worksheet`, `StreamSheet`, and `Elements::Workbook`, and expose `raw_values`, `formatted_values`, and `formulas` metadata on `Elements::Row` and `StreamRow`.
- Fallback serialization for DOM model attributes: Support serializing first-class `ws.state`, `ws.comments`, `ws.hyperlinks`, and `workbook.defined_names` in `Xlsxrb.write(workbook)` when not explicitly packaged inside `unmapped_data[:facade]` metadata.

### Changed
- Optimize chunked streaming read performance: Hoist `</sheetData>` tag lookup outside row iteration, eliminate byte delimiter array allocations, and terminate entry inflation early once worksheet rows are fully parsed, restoring streaming read throughput faster than in-memory read while preserving $O(1)$ memory consumption.

### Fixed
- Fallback cached value for formula cells: Emit `<v>0</v>` instead of omitting `<v>` when a formula cell does not specify a cached value, ensuring compatibility with OpenXML parsers (such as Roo and LibreOffice) that discard formula cells lacking `<v>` elements.
- Shared formula cell parsing: Support self-closing formula tags (e.g. `<f t="shared" si="0"/>`) in `WorksheetParser` regex scanners, preventing cells containing shared formulas from being silently omitted.
- Fixed missing `dcterms:created` and `dcterms:modified` serialization in `WorkbookWriter` core properties (`docProps/core.xml`).
- Prevent `Encoding::CompatibilityError` in `Xlsxrb.read` and `Cfb::Reader.cfb?` when inspecting raw string buffers in UTF-8 encoding.

## [0.1.13] - 2026-09-25

### Added
- Non-destructive template modification: Support modifying existing XLSX templates while preserving unmapped parts, media, charts, drawings, and VBA macros via `Xlsxrb.modify`.
- StreamSheet structural metadata accessors: Expose `merged_cells`, `auto_filter`, `data_validations`, and `conditional_formats` on `StreamSheet`.
- Pre-Push Smart Change Detection: Optimized `bin/pre-push --all` to automatically skip runtime RBS validation and pure logic mutation testing when unrelated files are modified, adding `--last` and `--force` controls.

### Changed
- True O(1) memory streaming read via chunked entry inflation in `ZipReader` and `WorksheetParser`, preventing unbounded memory retention on large worksheets.
- Extracted pure DSL normalizers (`Xlsxrb::DslHelpers`) and fast XML unescape utilities (`Ooxml::XmlBuilder.unescape`).
- Constrained `rexml` runtime dependency to `~> 3.0` and excluded documentation, build assets, and developer configurations from the gem package.

### Removed
- OpenTelemetry runtime dependency and tracer spans.

## [0.1.12] - 2026-09-20

### Added
- Mutation Testing Framework (`mutant`):
  - Full-scale integration of mutant test suite targeting the Functional Core (10 classes and modules, 38 subjects).
  - Achieved 100.00% Kill Rate (1,091 / 1,091 mutations killed, 0 alive) without exclusions.
  - Added dedicated root configuration [`.mutant.yml`](.mutant.yml) for streamlined CLI execution (`bundle exec mutant run -- '...'`).
  - Integrated `bundle exec rake mutant:pure` into GitHub Actions CI pipeline (`main.yml`), guaranteeing mutation-tested quality on every push and pull request.
  - Added architecture and testing guide in [`docs/MUTATION_TESTING.md`](docs/MUTATION_TESTING.md).
- Pre-Push Quality Verification Script (`bin/pre-push`): Added automated pre-push verification script running RuboCop, Steep type checking, RBS sync validation, Unit & Contract tests, RBS runtime type validation, and pure logic mutation testing, with automatic hook installation in `bin/setup`.

### Changed
- Functional Core & Imperative Shell Architecture:
  - Purified core business logic and domain entities into deterministic, side-effect-free class methods and predicates:
    - Extracted coordinate and validation predicates: `Elements::Cell.valid_value?`, `Cell.valid_coordinates?`, `Cell.calculate_column_letter`, `Cell.calculate_column_index`.
    - Extracted domain validation predicates: `Elements::Row.valid_index?`, `Elements::Column.valid_index?`, `Elements::Worksheet.valid_name?`.
    - Replaced date/time serial calculations in `Ooxml::Utils` with pure mathematical conversions (Julian Day math for `date_to_serial`, `divmod` arithmetic for `serial_to_datetime`, and 1900 leap year bug handling).
    - Extracted binary packing utilities in `Ooxml::ZipGenerator` (`le16`, `le32`, and bitfield arithmetic for `dos_datetime`).
    - Extracted pure XML character escaping logic in `Ooxml::XmlBuilder.escape`.
  - Modularized `Ooxml::Writer` into domain-specific mixin modules (`DrawingXml`, `FeaturesXml`, `StylesXml`).
  - Extracted shared DSL parameter normalization and range expansion into `Xlsxrb::DslHelpers`.
  - Modularized reader SAX listeners into domain-specific parser components.

### Fixed
- Boundary Validation & Error Reporting:
  - Strictly validate row number digits in `Elements::Cell.parse_ref` to reject malformed cell coordinate strings (e.g. `A0`, `A01`).
  - Harmonized supported value type checks and nil/empty handling across `Elements::Cell.validate`.
  - Corrected maximum row index boundary error messages in `Elements::Row.validate`.
  - Hardened cryptographic stream boundary checks and corrupted header handling in Standard and Agile encryption modes.
- RBS Runtime Type Validation:
  - Expanded `Elements::Workbook#initialize`, `#sheet`, and `[]` type annotations to accept `StreamSheet` instances used by streaming readers.
  - Made update block parameter optional in `Elements::Workbook#update_sheet` to allow testing missing block validation.
  - Supported `Numeric` types (Float and Integer) in `ChartBuilder` and `SeriesBuilder` kwargs (`width`, etc.).
  - Handled `RBS::Test::Tester::TypeError` in unit tests verifying block return type validation.

## [0.1.11] - 2026-08-19

### Performance
- Sub-Second Streaming Write (< 1.0s / 1,000,000 cells):
  - Achieved sub-second streaming write performance (0.97 s median for 1M cells) while maintaining 100% full SST (Shared String Table) XML deduplication and ISO/IEC 29500 compatibility.
  - Implemented 1-Pass Direct ZIP streaming pipeline, eliminating all intermediate tempfiles, disk seek-backs, and extra deflate flush cycles.
  - Added dedicated unstyled fast-path in `WorksheetWriter#write_row_values` with precomputed coordinate and integer lookup tables.
  - Pre-registered default date/time formatting styles, completely eliminating redundant per-row scan loops.
  - Streamlined `ZipWriter` instance variable lookups to eliminate per-chunk hash overhead.
  - Added `alias << row` to `WorksheetProxy` and `StreamWriter` for idiomatic Ruby streaming.
- High-Speed Streaming Read (1.61s / 1,000,000 cells):
  - Reduced streaming read execution time from 3.40 s down to 1.61 s (over 2x speedup).
  - Refactored `Elements::Cell` into an optimized lightweight class with `Cell.fast_create`, reducing 1M cell instantiation overhead by 5.2x (0.92s → 0.17s) while maintaining full pattern matching (`deconstruct`, `deconstruct_keys`) and immutability.
  - Replaced per-byte Ruby scanning in `WorksheetParser.fast_scan_cells_direct` with an Onigmo C-level regular expression scanner (`CELL_FAST_RE`), eliminating 100,000 intermediate XML substring allocations and nested byte loops.
  - Optimized `StreamRow#cells` to populate cell arrays via direct block traversal, eliminating Enumerator object allocations.

### Documentation
- Updated benchmark results, performance comparisons, and linear-scale SVG charts across `README.md` and `docs/PEER_LIBRARIES.md`.

## [0.1.10] - 2026-08-19

### Added
- Peer Libraries Ecosystem Guide ([docs/PEER_LIBRARIES.md](docs/PEER_LIBRARIES.md)): Introduced a respectful overview of the Ruby XLSX ecosystem featuring official self-descriptions, architectural tradeoffs (SST vs. Inline Strings, Streaming vs. In-Memory), and reproducible benchmarks across 9 popular Ruby XLSX gems.
- Visual Assets & Screen Previews:
  - Embedded interactive WebAssembly Playground live demo preview in `README.md`.
  - Added real-world Ruby LSP autocompletion and RBS type hint preview in `README.md`.
  - Created accurate, neutral linear-scale SVG benchmark performance chart.
- Multi-Layered Test Suite Expansion:
  - ECMA-376 XSD Schema Validation: XML schema validation suite ensuring strict element ordering and ISO/IEC 29500 compliance.
  - Contract Testing Suite: Parity verification between Streaming (`Xlsxrb.write`) and In-Memory (`Xlsxrb.build`) APIs.
  - Property-Based Testing (PBT): Expanded automated random generation tests for Row/Column invariants, styles, and edge cases.
  - Visual Regression Testing (VRT): Added new visual baselines for table styles, drawing shapes, and pivot tables.
  - E2E Interoperability Suite: Added tests for namespace-prefixed XML streaming, conditional formatting, and table structures.

### Fixed
- WebAssembly (ruby.wasm) Compatibility: Bundled `pp` and `prettyprint` standard libraries in `ruby.wasm` package to resolve REXML LoadError during browser-based evaluation.

### Changed
- Streamlined README: Refactored README from 363 to 175 lines, focusing on core motivation, clean 4-column feature matrix, concise usage examples, and direct links to specialized documentation.

## [0.1.9] - 2026-08-18

### Added
- Password Protection & Document Encryption ([MS-OFFCRYPTO] / [MS-CFB]): Full native Pure-Ruby support for reading, writing, and modifying password-protected Excel spreadsheets without any external C-extension dependencies.
  - Standard Encryption: AES-128-ECB and SHA-1 Key Derivation with CryptoAPI 50,000-spin hashing, fully interoperable across Microsoft Excel, LibreOffice, and Google Sheets.
  - Agile Encryption: Modern AES-256-CBC, PBKDF2/SHA-512, and HMAC-SHA512 data integrity verification.
  - Compound File Binary (CFB) Engine: Pure-Ruby reader and writer for OLE structured storage containers with Mini Stream, FAT/MiniFAT sectors, and Red-Black tree directory management.
  - Transparent Public API Integration: Added `password:` and `encryption_mode:` arguments to `Xlsxrb.read`, `Xlsxrb.write`, and `Xlsxrb.modify`.
  - Security & Threat Model Hardening:
    - Constant-time hash verification via `OpenSSL.secure_compare` to prevent timing attacks (CWE-208).
    - CSPRNG-backed salt, IV, and session key generation via `SecureRandom` (CWE-330).
    - DoS defense: spinCount limit ($\le 10\text{M}$), CFB circular sector chain loop detection in directory/FAT parsing, and `total_size` bounds validation (CWE-400, CWE-835).
    - Strict exception hierarchy (`EncryptedFileError`, `InvalidPasswordError`, `DecryptionError`).
  - Cross-Platform & Interoperability Validation: Bidirectional validation with Microsoft .NET OpenXML SDK and LibreOffice Calc.
  - WebAssembly (ruby.wasm) Support: Pre-packaged `docs/wasm/ruby.wasm` updated with document encryption support for browser playground.

## [0.1.8] - 2026-08-18

### Changed
- Unified Symmetric Entrypoints: Consolidated reading into `Xlsxrb.read` (supporting file path, IO, and raw binary string) and writing into `Xlsxrb.write` (supporting streaming blocks or in-memory Workbooks). Removed legacy `open`, `foreach`, and `generate` methods.
- Streaming-First Defaults: `Xlsxrb.read` yields and returns lightweight `StreamSheet` instances with $O(1)$ constant-memory consumption by default.
- Explicit In-Memory Materialization (`#load`): Stripped accidental random-access memory traps from `StreamSheet`; introduced explicit `StreamSheet#load` / `Workbook#load` (inspired by ActiveRecord Relations) to transition from lazy streaming to in-memory `Elements::Worksheet` / `Elements::Workbook`.
- `CoordinateAccess` Module: Extracted coordinate lookup methods (`[]`, `cell_value`, `row_at`, `first_row`, `last_row`, `cells`, `cells_hash`) into a dedicated `Xlsxrb::Elements::CoordinateAccess` mixin module included in `Elements::Worksheet`.

### Added
- Default Cell Streaming (`Xlsxrb::StreamRow`): Enabled streaming along both row and cell dimensions via `row.each_cell` and `sheet.each_cell`, parsing cells on-demand to handle sheets with thousands of columns in $O(1)$ constant memory.

## [0.1.7] - 2026-08-16

### Added
- Bundled native Ruby LSP Add-on (`RubyLsp::Xlsxrb::Addon`) for zero-configuration, context-aware method autocompletion and rich markdown documentation in VS Code and LSP-enabled editors for block arguments (`wb.`, `s.`, `sheet.`, `stream_writer.`, `stream_sheet.`).
- YARD documentation (`@param`, `@return`, `@example`) across all public APIs, builders, proxies, and elements.

### Developer Experience
- Enhanced `rbs-inline` type signatures across all facade methods and builder objects with automated RBS generation.
- Added VS Code Steep extension configuration for Dev Containers.

## [0.1.6] - 2026-08-15

### Added
- Support for unseekable IO streams (pipes, `ActionDispatch::Response::Buffer`, socket streams) across streaming reader and writer.
- Complete `rbs-inline` type annotations across all library source files and automated `sig/generated/` synchronization in CI.
- Automated Steep static type-checking workflow in GitHub Actions.
- Hash-compatible symbol indexing in `Elements::Row#[]` (`:cells`, `:index`, `:height`, `:attrs`) and `Elements::Cell#[]` (`:value`, `:ref`, `:style_index`).

### Performance
- Streaming Write: Replaced per-cell micro IO calls with row-level string buffering and fast-path serialization for unstyled cells (1,000,000 cells in 1.62s with standard SST).
- In-Memory Write: Optimized DOM serialization (`WorksheetWriter#write_row`) with row-level buffer aggregation, boosting 1,000,000 cells write from 22.33s to 4.06s (5.5x faster) and reducing GC count by 83%.
- Streaming Read: Implemented zero-allocation byte scanning for cell attributes (`r="..."`, `t="..."`, `s="..."`) and direct integer conversion for SST indexes, cutting GC count by ~70% (129 -> 40) and boosting 1,000,000 cells read to 3.38s.
- In-Memory Read: Reduced peak memory footprint by 57% (582 MB -> 250 MB).

### Documentation
- Overhauled `benchmark.rb` with `bundler/inline` for deterministic, zero-setup benchmark reproduction pinning peer gem versions.
- Updated README.md Motivation and Benchmark tables with verified default vs optional string storage architecture (SST vs Inline) and latest 1,000,000 cells measurements.

## [0.1.5] - 2026-08-11

### Added
- Functional API for updating existing files (`Xlsxrb.modify`).
- New `Workbook#update_sheet` and `Worksheet#update_cell` helpers for immutable data structures.
- Syntactic sugar for `styles:` in `sheet.row` using `Hash` and `Range` keys (e.g. `styles: { 0..4 => "header" }`).
- Support for inline anonymous styles as Hash objects (e.g. `styles: { 0 => { font: { bold: true } } }`).
- Support for `Range` and `Array` arguments in `sheet.column` to modify multiple columns at once.
- `WorkbookBuilder#[]` and `StreamWriter#[]` aliased to `#sheet` for elegant context switching.
- Block-based configuration for `font` and `border` properties inside `StyleBuilder`.
- Extensive inline RBS typing with strict generic types (eliminated `untyped` from the public API).
- Full `RBS::Test` runtime type validation enabled for the entire test suite.
- Extensive Excel limit warnings documented via YARD tags.
- Formal SemVer API contract with `@api public` tags for user-facing methods.
- Mutation testing (Mutant) and test coverage (SimpleCov) integrations.

### Changed
- Replaced `method_missing` with statically defined, fully typed methods in `WorksheetProxy`, `ChartBuilder`, and `SeriesBuilder`.
- Removed deprecated `instance_eval` block context for builders; explicit block arguments are now required.
- Hardened security and bounds checking across the DSL for `strict_excel_mode`.
- Fixed numerous Rubocop linting violations and standardized style rules.
- Add specification reference policy and implemented specification mapping (`docs/SPEC_SOURCES.md`).
- Introduce unified event-based streaming and parsing architecture (`Ooxml::Event` and event streams for WorksheetParser/SharedStringsParser).
- Restructure test suite into four distinct tiers (Unit, Contract, E2E, Visual) and renamed `facade_test.rb` to `public_api_test.rb`.
- Add visual examples gallery (`examples/visual/` and `docs/visual/README.md`) and promoted it via animated GIF in the main README.
- Implement visual regression testing (VRT) pipeline comparing generated sheets against reference baselines.

## [0.1.0] - 2026-03-25

- Initial release
