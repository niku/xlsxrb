## [Unreleased]
### Added
- Support for unseekable IO streams (pipes, `ActionDispatch::Response::Buffer`, socket streams) across streaming reader and writer.
- Complete `rbs-inline` type annotations across all library source files and automated `sig/generated/` synchronization in CI.
- Automated Steep static type-checking workflow in GitHub Actions.
- Hash-compatible symbol indexing in `Elements::Row#[]` (`:cells`, `:index`, `:height`, `:attrs`) and `Elements::Cell#[]` (`:value`, `:ref`, `:style_index`).

### Performance
- **Streaming Write**: Replaced per-cell micro IO calls with row-level string buffering and fast-path serialization for unstyled cells (1,000,000 cells in 1.62s with standard SST).
- **In-Memory Write**: Optimized DOM serialization (`WorksheetWriter#write_row`) with row-level buffer aggregation, boosting 1,000,000 cells write from 22.33s to 4.06s (5.5x faster) and reducing GC count by 83%.
- **Streaming Read**: Implemented zero-allocation byte scanning for cell attributes (`r="..."`, `t="..."`, `s="..."`) and direct integer conversion for SST indexes, cutting GC count by ~70% (129 -> 40) and boosting 1,000,000 cells read to 3.38s.
- **In-Memory Read**: Reduced peak memory footprint by 57% (582 MB -> 250 MB).

### Documentation
- Overhauled `benchmark.rb` with `bundler/inline` for deterministic, zero-setup benchmark reproduction pinning peer gem versions.
- Updated README.md Motivation and Benchmark tables with verified default vs optional string storage architecture (SST vs Inline) and latest 1,000,000 cells measurements.

## [0.1.5] - 2026-08-11

## [0.1.5] - 2026-08-11

## [0.1.5] - 2026-08-11

## [0.1.5] - 2026-08-11

## [0.1.5] - 2026-08-11

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
- Comprehensive mutation testing (Mutant) and test coverage (SimpleCov) integrations.

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
