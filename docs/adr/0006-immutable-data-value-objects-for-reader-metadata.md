# ADR 0006: Immutable Data Value Objects for Reader Metadata and Collections

* Status: Accepted
* Date: 2026-11-06
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::Elements::MetadataRecord`](../../lib/xlsxrb/elements/metadata.rb), [`Xlsxrb::Elements`](../../lib/xlsxrb/elements.rb), [`Xlsxrb::StreamSheet`](../../lib/xlsxrb/stream_sheet.rb), [`Xlsxrb::Elements::Styles`](../../lib/xlsxrb/elements/styles.rb)

---

## Context

In xlsxrb 0.1, Reader metadata inspection returned raw `Hash`es:
* Singular sheet settings (`sheet.page_setup`, `sheet.page_margins`, `sheet.freeze_pane`, `sheet.sheet_properties`, `wb.core_properties`, etc.) returned untyped Hashes with varying key conventions (`:col` vs `:column`, `:sz` vs `:size`).
* Collection properties returned raw Hashes or mixed types. For example, `sheet.hyperlinks` returned a raw Hash of Hashes (`{ "A1" => { url: "...", tooltip: "..." } }`), whereas `sheet.hyperlink("A1")` and `cell.hyperlink` returned an `Elements::Hyperlink` object.
* Styles collections (`Styles#fonts`, `Styles#fills`, `Styles#borders`, `Styles#cell_xfs`) returned raw hashes, preventing compile-time RBS type checking and autocompletion in language servers (Ruby LSP).

This design lacked type safety, forced callers into defensive `dig` / `is_a?(Hash)` checks, and hindered thread and Ractor safety because mutable Hashes could be modified after parsing.

---

## Guiding Principles

1. Immutable Value Objects by Default:
   Domain metadata parsed from OpenXML documents is immutable. Ruby 3.2+ `Data.define` value objects provide memory efficiency, constant shareability across Ractors, and descriptive equality semantics.
2. Symmetrical Type Representation:
   Singular lookups and collection lookups must yield the exact same type (e.g. `sheet.hyperlink("A1")` and values of `sheet.hyperlinks` both return `Elements::Hyperlink`).
3. Seamless Hash Backward Compatibility:
   All metadata value objects implement `MetadataRecord` (`#[](key)`, `#dig`, `#fetch`, `#merge`, `#to_h`, `#to_hash`, `#empty?`, and `#==(other)`) so existing code expecting Hash access functions without modification.
4. Writer Acceptance Symmetry:
   Writers accept both `Elements::*` `Data` value objects and raw `Hash`es or kwargs interchangeably.

---

## Decision

### 1. 1-to-1 Metadata Properties Return Elements::<Name>
On `Elements::Workbook`, `Elements::Worksheet`, `StreamSheet`, and `Elements::Styles`:
* `core_properties`: `Elements::CoreProperties`
* `app_properties`: `Elements::AppProperties`
* `workbook_properties`: `Elements::WorkbookProperties`
* `workbook_protection`: `Elements::WorkbookProtection?`
* `page_setup`: `Elements::PageSetup`
* `page_margins`: `Elements::PageMargins?`
* `header_footer`: `Elements::HeaderFooter`
* `print_options`: `Elements::PrintOptions`
* `print_titles`: `Elements::PrintTitles?`
* `freeze_pane`: `Elements::FreezePane?`
* `split_pane`: `Elements::SplitPane?`
* `selection`: `Elements::Selection?`
* `sheet_protection`: `Elements::SheetProtection?`
* `sheet_properties`: `Elements::SheetProperties`
* `sheet_format`: `Elements::SheetFormat`
* `sheet_view`: `Elements::SheetView`
* `phonetic_properties`: `Elements::PhoneticProperties?`
* `sheet_calc_properties`: `Elements::SheetCalcProperties?`
* `sort_state`: `Elements::SortState?`
* `background_image`: `Elements::BackgroundImage?`
* `table_styles`: `Elements::TableStyles`
* `colors`: `Elements::Colors`

### 2. Dynamic-Keyed Maps Return Hash[Key, Elements::<Name>]
Collections with dynamic, user-defined, or coordinate keys return frozen Hashes of value objects:
* `sheet.hyperlinks`: `Hash[String, Elements::Hyperlink]` (keyed by cell reference like `"A1"`)
* `sheet.filter_columns`: `Hash[Integer, Elements::FilterColumn]` (keyed by 0-based column index)
* `workbook.custom_properties`: `Hash[String, Elements::CustomProperty]` (keyed by property name)
* `styles.number_formats`: `Hash[Integer, String]` (keyed by numFmtId)

### 3. Indexed Collections Return Array[Elements::<Name>]
* `workbook.defined_names`: `Array[Elements::DefinedName]`
* `sheet.comments`: `Array[Elements::Comment]`
* `sheet.threaded_comments`: `Array[Elements::ThreadedComment]`
* `sheet.data_validations`: `Array[Elements::DataValidation]`
* `sheet.conditional_formats`: `Array[Elements::ConditionalFormat]`
* `sheet.tables`: `Array[Elements::Table]`
* `sheet.row_breaks`: `Array[Elements::PageBreak]`
* `sheet.column_breaks`: `Array[Elements::PageBreak]`
* `sheet.sparkline_groups`: `Array[Elements::SparklineGroup]`
* `sheet.charts`: `Array[Elements::Chart]`
* `sheet.shapes`: `Array[Elements::Shape]`
* `sheet.pivot_tables`: `Array[Elements::PivotTable]`
* `styles.fonts`: `Array[Elements::Font]`
* `styles.fills`: `Array[Elements::Fill]`
* `styles.borders`: `Array[Elements::Border]`
* `styles.cell_xfs`: `Array[Elements::CellXf]`
* `styles.dxfs`: `Array[Elements::Dxf]`
* `styles.cell_styles`: `Array[Elements::CellStyle]`

### 4. MetadataRecord Compatibility Mixin and Deprecation Schedule
Every `Data.define` class includes `Elements::MetadataRecord`:
* In 0.1.x, hash-like key access (`obj[:key]`, `obj["key"]`, `#key?`, `#dig`, `#fetch`, `#merge`) and Hash equality (`obj == { ... }`) were supported with deprecation warnings.
* Complete removal in 0.2.0: In 0.2.0, the hash compatibility methods (`#[]`, `#key?`, `#dig`, `#fetch`, `#merge`, and equality against Hashes) have been removed, leaving `Elements::*` value objects accessed exclusively through method calls.
* Serialization methods (`#to_h` and `#to_hash`) and `#empty?` remain permanently supported on `MetadataRecord` for dictionary serialization, empty checks, and keyword splatting (`**obj`).

### 5. Catch-All Options Attribute for Extensible Elements
In the ECMA-376 OpenXML specification, elements such as `<c:chart>`, `<xdr:sp>` (shapes), and `<cfRule>` (conditional formatting rules) do not define any XML tag or attribute named `options`. Instead, they define open-ended hierarchies of drawing, fill, line, and layout sub-elements.

To preserve arbitrary formatting and styling options passed via Writer DSLs (such as `fill_color:`, `font_color:`, line styles) without silent loss during `Elements::*` instantiation:
* Extensible elements (`Elements::Chart`, `Elements::Shape`, `Elements::ConditionalFormat`, `Elements::ConditionalFormatRule`, `Elements::PivotTable`) include an `options` catch-all attribute.
* Unknown keyword arguments (`**extra`) are captured in the `options` Hash.
* Attribute resolution precedence: explicit `Data.define` members always take precedence. If a key is not a declared member, `MetadataRecord#[]` and `MetadataRecord#key?` fall back to inspecting `options`.
* Serialization transparency: `MetadataRecord#to_h` unrolls `options` into the top-level Hash so that OpenXML serializers and writer methods receive a flat dictionary without requiring callers to manually drill into nested options containers.

### 6. Elements::Workbook Metadata Promotion and Styles Normalization
* `Elements::Workbook` promotes `:workbook_properties`, `:core_properties`, `:app_properties`, `:workbook_protection`, and `:custom_properties` to direct `Data.define` members.
* Normalizes `styles` to `Elements::Styles` upon initialization.
* Supports `Symbol` sheet names in `#sheet`, `#[]`, and `#defined_name`.

### 7. Elements::Cell Value Object Normalization
* `Elements::Cell` normalizes string and hash formula inputs to `Elements::Formula` value objects via `Elements::Cell.normalize_formula`.
* Normalizes string and hash comment inputs to `Elements::Comment` value objects.

### 8. StreamSheet Column Indexing and Fast-Path Inspection
* `StreamSheet#first_column` and `#last_column` are deprecated in favor of 0-based `#column_range` and `#dimension_bounds`.
* Implements tag-based fast paths in `merged_cells`, `data_validations`, and `conditional_formats` that inspect entry presence before scanning XML buffers.

---

## Consequences

### Positive
* Type Safety and Autocompletion: Editors and type checkers resolve methods directly on `sheet.page_setup.orientation` without union type ambiguities.
* Zero Discrepancy: `sheet.hyperlinks["A1"]` and `sheet.hyperlink("A1")` return the exact same `Elements::Hyperlink` type.
* Ractor Isolation: Value objects are deeply frozen and constant shareable across Ractors.
* Safe Transitional Compatibility: Existing call sites using `sheet.page_setup[:orientation]` continue to work with deprecation warnings in 0.1.x, providing a smooth migration path before complete removal in 0.2.0.

### Negative / Trade-offs
* Allocation Profile: Instantiating lightweight `Data` objects rather than reusing raw parser hashes. Microbenchmarks verify this has negligible impact on parse throughput while keeping memory consumption $O(1)$ for streaming pipelines.
