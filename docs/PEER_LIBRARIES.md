# Ruby XLSX Ecosystem & Peer Libraries

The Ruby ecosystem includes several XLSX libraries, each designed with specific architectural tradeoffs for different problem spaces.

This document provides an overview of peer libraries in the Ruby ecosystem, explains architectural tradeoffs (such as Streaming vs. In-Memory and Shared String Tables vs. Inline Strings), and shares benchmark measurements.

---

## The Peer Libraries

| Library | I/O | Official Self-Description / Focus | Best Fit (In Our View) |
| :--- | :---: | :--- | :--- |
| [roo](https://rubygems.org/gems/roo) | `R` | *"Roo can access the contents of various spreadsheet files (Excelx, LibreOffice, OpenOffice, CSV)."* | Unified interface for reading across diverse spreadsheet formats. |
| [creek](https://rubygems.org/gems/creek) | `R` | *"A Ruby gem that streams and parses large Excel (xlsx and xlsm) files fast and efficiently."* | Streaming large spreadsheet uploads row-by-row with lightweight SAX parsing. |
| [xsv](https://rubygems.org/gems/xsv) | `R` | *"A fast and lightweight xlsx parser that provides nothing a CSV parser wouldn't."* | High-speed, CSV-like tabular data ingestion without styling overhead. |
| [simple_xlsx_reader](https://rubygems.org/gems/simple_xlsx_reader) | `R` | *"Read xlsx data the Ruby way"* — parses sheets into Ruby primitives with low memory. | Memory-conscious tabular data extraction directly into Ruby types. |
| [caxlsx / axlsx](https://rubygems.org/gems/caxlsx) | `W` | *"Excel OOXML (xlsx) with charts, styles, images and autowidth columns"* with full schema validation. | Generating rich, styled business reports with charts, images, and visual design. |
| [write_xlsx](https://rubygems.org/gems/write_xlsx) | `W` | Pure Ruby port of Perl's `Excel::Writer::XLSX` to create files in modern Excel 2007+ format. | Creating complex spreadsheets requiring exact Excel feature parity. |
| [xlsxtream](https://rubygems.org/gems/xlsxtream) | `W` | *"A streaming XLSX spreadsheet writer"* allowing very efficient writing of CSV-style data. | Low-memory streaming exports of tabular datasets. |
| [fast_excel](https://rubygems.org/gems/fast_excel) | `W` | *"Ultra Fast Excel Writer"* — C-extension wrapper for `libxlsxwriter` with constant memory mode. | High-throughput spreadsheet generation when C-extensions are available. |
| [rubyXL](https://rubygems.org/gems/rubyXL) | `RW` | *"Allows the parsing, creation, and manipulation of Microsoft Excel (.xlsx/.xlsm) Documents."* | Full document DOM inspection, in-memory cell modification, and template editing. |
| [xlsxrb](https://github.com/niku/xlsxrb) | `RW` | Pure Ruby library unifying streaming read/write ($O(1)$ memory) and in-memory manipulation with native encryption. | Unified reading, writing, template modification, and password encryption in pure Ruby. |


---

## Architectural Tradeoffs

Spreadsheet libraries must balance multiple competing dimensions: memory consumption, execution speed, formatting capabilities, and strict specification compliance.

### 1. Memory & Execution Model: Streaming vs. In-Memory

```
┌─────────────────────────────────────────────────────────────┐
│                      Execution Models                       │
├──────────────────────────────┬──────────────────────────────┤
│       Streaming Model        │       In-Memory Model        │
├──────────────────────────────┼──────────────────────────────┤
│ • Processes rows on-the-fly  │ • Builds complete DOM tree   │
│ • O(1) constant RAM footprint│ • Enables random cell access │
│ • Cannot seek backward       │ • High RAM on large datasets │
│ • Ideal for batch exports/ETL│ • Ideal for templates/edits  │
└──────────────────────────────┴──────────────────────────────┘
```

* Streaming Model (`xlsxrb`, `xlsxtream`, `simple_xlsx_reader`, `roo`, `creek`, `xsv`):
  Rows and cells are processed sequentially and flushed/discarded immediately. This keeps memory usage constant and predictable, regardless of whether the file has 10 rows or 1,000,000 rows. However, random access (e.g., modifying `cell("A1")` after writing row 100) is not possible.
* In-Memory Model (`xlsxrb`, `caxlsx`, `write_xlsx`, `rubyXL`):
  The entire workbook structure is parsed into Ruby objects, providing flexibility to inspect, modify, insert, or reorder cells and worksheets. The tradeoff is that memory consumption scales with the number of cells.

---

### 2. String Storage Architecture: Shared String Table (SST) vs. Inline Strings

The OpenXML (ECMA-376) specification defines two ways to store text in cells:

```xml
<!-- 1. Shared String Table (SST): Deduplicated dictionary reference -->
<c r="A1" t="s"><v>0</v></c>

<!-- 2. Inline String: Raw text payload inside the cell -->
<c r="A1" t="inlineStr"><is><t>Hello World</t></is></c>
```

#### Shared String Table (SST)
* How it works: Strings across all worksheets are collected into a single central dictionary (`xl/sharedStrings.xml`). Cells only store numeric integer IDs pointing to dictionary entries.
* Strengths:
  * Smaller file footprint: Deduplication significantly reduces the uncompressed XML size (typically 50% to 80% smaller for business datasets with repetitive categories, statuses, dates, and labels).
  * Standard Microsoft Excel behavior: Excel defaults to SST. Opening SST-based spreadsheets in Excel consumes less memory and renders faster.
  * Rich Text & Shared Styles: Supports rich text formatting within strings.
* Tradeoff:
  * Writing requires managing a string table dictionary or making a multi-pass serialization, adding slight CPU overhead during generation.

#### Inline Strings
* How it works: Text is written directly into each `<c>` element (`<is><t>...</t></is>`) as the stream proceeds.
* Strengths:
  * Raw Throughput: Bypassing string deduplication allows immediate row-by-row flushing with minimal CPU overhead (as demonstrated by `xlsxtream`).
* Tradeoff:
  * Produces significantly larger raw XML files when strings repeat, and some third-party spreadsheet viewers or legacy tools have limited support for inline strings compared to SST.

---

## Detailed Benchmark Results

The following benchmarks evaluate processing 1,000,000 cells (100,000 rows × 10 columns) containing standard business data (integers, strings, floats, booleans, and dates) across 3 isolated subprocess runs.

Environment: Ruby 4.0.9 on aarch64-linux (the `ruby:4.0` Docker image on Apple silicon), measured on 2026-11-09. All libraries run with YJIT enabled (`RUBY_YJIT_ENABLE=1`), because Rails 7.2 and later enable YJIT by default (`config.yjit`). CRuby itself does not enable YJIT by default. xlsxrb writes with its default ZIP compression level 1 ([ADR 0009](adr/0009-default-zip-compression-level.md)); the other writers use zlib's default level 6, so their files are smaller.

### Write Performance (1,000,000 cells)

| Library | Version | Model | String Storage | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| xlsxrb (Streaming) | - | Streaming | SST (Shared) | 0.46 s | 0.47 s | 67.9 MB | 14 |
| xlsxtream | 3.1.0 | Streaming | Inline String | 1.04 s | 1.06 s | 19.5 MB | 1038 |
| fast_excel | 0.6.1 | Streaming | SST (Shared) | 1.61 s | 1.60 s | 151.0 MB | 211 |
| xlsxrb (In-Memory) | - | In-Memory | SST (Shared) | 1.86 s | 1.82 s | 379.3 MB | 13 |
| write_xlsx | 1.15.1 | In-Memory | SST (Shared) | 2.40 s | 2.41 s | 212.3 MB | 27 |
| caxlsx | 4.5.0 | In-Memory | Inline String | 2.66 s | 2.67 s | 183.5 MB | 23 |
| rubyXL | 3.4.38 | In-Memory | Inline String | 36.37 s | 36.03 s | 2183.0 MB | 103 |

### Read Performance (1,000,000 cells)

| Library | Version | Model | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| xlsxrb (Streaming) | - | Streaming | 1.74 s | 1.74 s | 73.2 MB | 90 |
| simple_xlsx_reader | 5.1.0 | Streaming | 3.82 s | 3.85 s | 32.5 MB | 1164 |
| xlsxrb (In-Memory) | - | In-Memory | 4.18 s | 4.22 s | 521.6 MB | 45 |
| creek | 2.6.3 | Streaming | 7.52 s | 7.55 s | 826.7 MB | 586 |
| roo | 4.0.0 | Streaming | 8.35 s | 8.44 s | 135.0 MB | 125 |
| xsv | 1.4.1 | Streaming | 13.61 s | 13.61 s | 66.5 MB | 817 |
| rubyXL | 3.4.38 | In-Memory | 34.41 s | 36.98 s | 2198.7 MB | 139 |

Peak memory of the xlsxrb in-memory model is higher than in the 2026-08-19 measurement (write 209.0 MB, read 304.3 MB), which was taken in a different environment without YJIT.

---

## Reproducing Benchmarks Locally

The benchmark suite leverages [`bundler/inline`](https://bundler.io/v2.5/guides/bundler_in_a_single_file_ruby_script.html) to run each library in an isolated subprocess (`Bundler.with_unbundled_env`), eliminating cross-gem pollution and ensuring accurate memory measurements.

To run the suite on your machine:

```bash
RUBY_YJIT_ENABLE=1 ruby benchmark.rb 100000 10
```

Omit `RUBY_YJIT_ENABLE=1` to measure the interpreter.
