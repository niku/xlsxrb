## [0.1.11] - 2026-08-19

### Performance
- **Sub-Second Streaming Write (< 1.0s / 1,000,000 cells)**:
  - Achieved sub-second streaming write performance (**0.97 s** median for 1M cells) while maintaining 100% full SST (Shared String Table) XML deduplication and ISO/IEC 29500 compatibility.
  - Implemented 1-Pass Direct ZIP streaming pipeline, eliminating all intermediate tempfiles, disk seek-backs, and extra deflate flush cycles.
  - Added dedicated unstyled fast-path in `WorksheetWriter#write_row_values` with precomputed coordinate and integer lookup tables.
  - Pre-registered default date/time formatting styles, completely eliminating redundant per-row scan loops.
  - Streamlined `ZipWriter` instance variable lookups to eliminate per-chunk hash overhead.
  - Added `alias << row` to `WorksheetProxy` and `StreamWriter` for idiomatic Ruby streaming.
- **High-Speed Streaming Read (1.61s / 1,000,000 cells)**:
  - Reduced streaming read execution time from **3.40 s** down to **1.61 s** (over 2x speedup).
  - Refactored `Elements::Cell` into an optimized lightweight class with `Cell.fast_create`, reducing 1M cell instantiation overhead by 5.2x (0.92s → 0.17s) while maintaining full pattern matching (`deconstruct`, `deconstruct_keys`) and immutability.
  - Replaced per-byte Ruby scanning in `WorksheetParser.fast_scan_cells_direct` with an Onigmo C-level regular expression scanner (`CELL_FAST_RE`), eliminating 100,000 intermediate XML substring allocations and nested byte loops.
  - Optimized `StreamRow#cells` to populate cell arrays via direct block traversal, eliminating Enumerator object allocations.

### Documentation
- Updated benchmark results, performance comparisons, and linear-scale SVG charts across `README.md` and `docs/PEER_LIBRARIES.md`.

## [0.1.10] - 2026-08-19

### Added
- **Peer Libraries Ecosystem Guide ([docs/PEER_LIBRARIES.md](docs/PEER_LIBRARIES.md))**: Introduced a respectful overview of the Ruby XLSX ecosystem featuring official self-descriptions, architectural tradeoffs (SST vs. Inline Strings, Streaming vs. In-Memory), and reproducible benchmarks across 9 popular Ruby XLSX gems.
- **Visual Assets & Screen Previews**:
  - Embedded interactive WebAssembly Playground live demo preview in `README.md`.
  - Added real-world Ruby LSP autocompletion and RBS type hint preview in `README.md`.
  - Created accurate, neutral linear-scale SVG benchmark performance chart.
- **Enterprise-Grade Test Suite Expansion**:
  - **ECMA-376 XSD Schema Validation**: Comprehensive XML schema validation suite ensuring strict element ordering and ISO/IEC 29500 compliance.
  - **Contract Testing Suite**: Comprehensive parity verification between Streaming (`Xlsxrb.write`) and In-Memory (`Xlsxrb.build`) APIs.
  - **Property-Based Testing (PBT)**: Expanded automated random generation tests for Row/Column invariants, styles, and edge cases.
  - **Visual Regression Testing (VRT)**: Added new visual baselines for table styles, drawing shapes, and pivot tables.
  - **E2E Interoperability Suite**: Added tests for namespace-prefixed XML streaming, conditional formatting, and table structures.

### Fixed
- **WebAssembly (ruby.wasm) Compatibility**: Bundled `pp` and `prettyprint` standard libraries in `ruby.wasm` package to resolve REXML LoadError during browser-based evaluation.

### Changed
- **Streamlined README**: Refactored README from 363 to 175 lines, focusing on core motivation, clean 4-column feature matrix, concise usage examples, and direct links to specialized documentation.

## [0.1.9] - 2026-08-18

### Added
- **Password Protection & Document Encryption ([MS-OFFCRYPTO] / [MS-CFB])**: Full native Pure-Ruby support for reading, writing, and modifying password-protected Excel spreadsheets without any external C-extension dependencies.
  - **Standard Encryption**: AES-128-ECB and SHA-1 Key Derivation with CryptoAPI 50,000-spin hashing, fully interoperable across Microsoft Excel, LibreOffice, and Google Sheets.
  - **Agile Encryption**: Modern AES-256-CBC, PBKDF2/SHA-512, and HMAC-SHA512 data integrity verification.
  - **Compound File Binary (CFB) Engine**: Pure-Ruby reader and writer for OLE structured storage containers with Mini Stream, FAT/MiniFAT sectors, and Red-Black tree directory management.
  - **Transparent Public API Integration**: Added `password:` and `encryption_mode:` arguments to `Xlsxrb.read`, `Xlsxrb.write`, and `Xlsxrb.modify`.
  - **Security & Threat Model Hardening**:
    - Constant-time hash verification via `OpenSSL.secure_compare` to prevent timing attacks (CWE-208).
    - CSPRNG-backed salt, IV, and session key generation via `SecureRandom` (CWE-330).
    - Robust DoS defense: spinCount limit ($\le 10\text{M}$), CFB circular sector chain loop detection in directory/FAT parsing, and `total_size` bounds validation (CWE-400, CWE-835).
    - Strict exception hierarchy (`EncryptedFileError`, `InvalidPasswordError`, `DecryptionError`).
  - **Cross-Platform & Interoperability Validation**: Bidirectional validation with Microsoft .NET OpenXML SDK and LibreOffice Calc.
  - **WebAssembly (ruby.wasm) Support**: Pre-packaged `docs/wasm/ruby.wasm` updated with document encryption support for browser playground.

## [0.1.8] - 2026-08-18

### Changed
- **Unified Symmetric Entrypoints**: Consolidated reading into `Xlsxrb.read` (supporting file path, IO, and raw binary string) and writing into `Xlsxrb.write` (supporting streaming blocks or in-memory Workbooks). Removed legacy `open`, `foreach`, and `generate` methods.
- **Streaming-First Defaults**: `Xlsxrb.read` yields and returns lightweight `StreamSheet` instances with $O(1)$ constant-memory consumption by default.
- **Explicit In-Memory Materialization (`#load`)**: Stripped accidental random-access memory traps from `StreamSheet`; introduced explicit `StreamSheet#load` / `Workbook#load` (inspired by ActiveRecord Relations) to transition from lazy streaming to in-memory `Elements::Worksheet` / `Elements::Workbook`.
- **`CoordinateAccess` Module**: Extracted coordinate lookup methods (`[]`, `cell_value`, `row_at`, `first_row`, `last_row`, `cells`, `cells_hash`) into a dedicated `Xlsxrb::Elements::CoordinateAccess` mixin module included in `Elements::Worksheet`.

### Added
- **Default Cell Streaming (`Xlsxrb::StreamRow`)**: Enabled streaming along both row and cell dimensions via `row.each_cell` and `sheet.each_cell`, parsing cells on-demand to handle sheets with thousands of columns in $O(1)$ constant memory.

## [0.1.7] - 2026-08-16

### Added
- Bundled native Ruby LSP Add-on (`RubyLsp::Xlsxrb::Addon`) for zero-configuration, context-aware method autocompletion and rich markdown documentation in VS Code and LSP-enabled editors for block arguments (`wb.`, `s.`, `sheet.`, `stream_writer.`, `stream_sheet.`).
- Comprehensive YARD documentation (`@param`, `@return`, `@example`) across all public APIs, builders, proxies, and elements.

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
- **Streaming Write**: Replaced per-cell micro IO calls with row-level string buffering and fast-path serialization for unstyled cells (1,000,000 cells in 1.62s with standard SST).
- **In-Memory Write**: Optimized DOM serialization (`WorksheetWriter#write_row`) with row-level buffer aggregation, boosting 1,000,000 cells write from 22.33s to 4.06s (5.5x faster) and reducing GC count by 83%.
- **Streaming Read**: Implemented zero-allocation byte scanning for cell attributes (`r="..."`, `t="..."`, `s="..."`) and direct integer conversion for SST indexes, cutting GC count by ~70% (129 -> 40) and boosting 1,000,000 cells read to 3.38s.
- **In-Memory Read**: Reduced peak memory footprint by 57% (582 MB -> 250 MB).

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
