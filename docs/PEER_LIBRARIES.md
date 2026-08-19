# Ruby XLSX Ecosystem & Peer Libraries

The Ruby ecosystem is fortunate to have a rich set of mature, well-engineered XLSX libraries. Each library represents deliberate architectural choices tailored for specific problem spaces.

This document provides a respectful overview of the peer libraries in the Ruby ecosystem, explains the underlying engineering tradeoffs (such as Streaming vs. In-Memory and Shared String Tables vs. Inline Strings), and shares comprehensive benchmark measurements.

---

## The Peer Libraries

| Library | I/O | Official Self-Description / Focus | Best Fit (In Our View) |
| :--- | :---: | :--- | :--- |
| **[roo](https://rubygems.org/gems/roo)** | `R` | *"Roo can access the contents of various spreadsheet files (Excelx, LibreOffice, OpenOffice, CSV)."* | Unified interface for reading across diverse spreadsheet formats. |
| **[creek](https://rubygems.org/gems/creek)** | `R` | *"A Ruby gem that streams and parses large Excel (xlsx and xlsm) files fast and efficiently."* | Streaming large spreadsheet uploads row-by-row with lightweight SAX parsing. |
| **[xsv](https://rubygems.org/gems/xsv)** | `R` | *"A fast and lightweight xlsx parser that provides nothing a CSV parser wouldn't."* | High-speed, CSV-like tabular data ingestion without styling overhead. |
| **[simple_xlsx_reader](https://rubygems.org/gems/simple_xlsx_reader)** | `R` | *"Read xlsx data the Ruby way"* — parses sheets into Ruby primitives with low memory. | Memory-conscious tabular data extraction directly into Ruby types. |
| **[caxlsx / axlsx](https://rubygems.org/gems/caxlsx)** | `W` | *"Excel OOXML (xlsx) with charts, styles, images and autowidth columns"* with full schema validation. | Generating rich, styled business reports with charts, images, and visual design. |
| **[write_xlsx](https://rubygems.org/gems/write_xlsx)** | `W` | Pure Ruby port of Perl's `Excel::Writer::XLSX` to create files in modern Excel 2007+ format. | Creating complex spreadsheets requiring exact Excel feature parity. |
| **[xlsxtream](https://rubygems.org/gems/xlsxtream)** | `W` | *"A streaming XLSX spreadsheet writer"* allowing very efficient writing of CSV-style data. | Ultra-fast, low-memory streaming exports of massive tabular datasets. |
| **[fast_excel](https://rubygems.org/gems/fast_excel)** | `W` | *"Ultra Fast Excel Writer"* — C-extension wrapper for `libxlsxwriter` with constant memory mode. | Maximum-throughput spreadsheet generation when C-extensions are available. |
| **[rubyXL](https://rubygems.org/gems/rubyXL)** | `RW` | *"Allows the parsing, creation, and manipulation of Microsoft Excel (.xlsx/.xlsm) Documents."* | Full document DOM inspection, in-memory cell modification, and template editing. |
| **[xlsxrb](https://github.com/niku/xlsxrb)** | `RW` | Pure Ruby library unifying streaming read/write ($O(1)$ memory) and in-memory manipulation with native encryption. | Unified reading, writing, template modification, and password encryption in pure Ruby. |


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

* **Streaming Model** (`xlsxrb`, `xlsxtream`, `simple_xlsx_reader`, `roo`, `creek`, `xsv`):
  Rows and cells are processed sequentially and flushed/discarded immediately. This keeps memory usage completely flat and predictable, regardless of whether the file has 10 rows or 1,000,000 rows. However, random access (e.g., modifying `cell("A1")` after writing row 100) is not possible.
* **In-Memory Model** (`xlsxrb`, `caxlsx`, `write_xlsx`, `rubyXL`):
  The entire workbook structure is parsed into Ruby objects, providing complete flexibility to inspect, modify, insert, or reorder cells and worksheets. The tradeoff is that memory consumption scales with the number of cells.

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
* **How it works**: Strings across all worksheets are collected into a single central dictionary (`xl/sharedStrings.xml`). Cells only store numeric integer IDs pointing to dictionary entries.
* **Strengths**:
  * **Smaller file footprint**: Deduplication significantly reduces the uncompressed XML size (typically 50% to 80% smaller for business datasets with repetitive categories, statuses, dates, and labels).
  * **Standard Microsoft Excel behavior**: Excel defaults to SST. Opening SST-based spreadsheets in Excel consumes less memory and renders faster.
  * **Rich Text & Shared Styles**: Supports rich text formatting within strings.
* **Tradeoff**:
  * Writing requires managing a string table dictionary or making a multi-pass serialization, adding slight CPU overhead during generation.

#### Inline Strings
* **How it works**: Text is written directly into each `<c>` element (`<is><t>...</t></is>`) as the stream proceeds.
* **Strengths**:
  * **Raw Throughput**: Bypassing string deduplication allows immediate row-by-row flushing with minimal CPU overhead (as demonstrated by `xlsxtream`).
* **Tradeoff**:
  * Produces significantly larger raw XML files when strings repeat, and some third-party spreadsheet viewers or legacy tools have limited support for inline strings compared to SST.

---

## Detailed Benchmark Results

The following benchmarks evaluate processing **1,000,000 cells** (100,000 rows × 10 columns) containing standard business data (integers, strings, floats, booleans, and dates) across 3 isolated subprocess runs.

### Write Performance (1,000,000 cells)

| Library | Version | Model | String Storage | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **xlsxtream** | 3.1.0 | Streaming | Inline String | **1.19 s** | 1.20 s | **18.2 MB** | 1072.0 |
| **xlsxrb (Streaming)** | - | Streaming | SST (Shared) | **1.73 s** | 1.65 s | 94.4 MB | 39.0 |
| **fast_excel** | 0.5.0 | Streaming | SST (Shared) | 1.89 s | 1.89 s | 148.2 MB | 245.0 |
| **xlsxrb (In-Memory)** | - | In-Memory | SST (Shared) | 3.84 s | 3.83 s | 278.3 MB | 32.0 |
| **write_xlsx** | 1.15.0 | In-Memory | SST (Shared) | 4.32 s | 4.34 s | 201.2 MB | 33.0 |
| **caxlsx** | 4.5.0 | In-Memory | Inline String | 5.15 s | 5.12 s | 188.6 MB | 23.0 |
| **rubyXL** | 3.4.38 | In-Memory | Inline String | 38.81 s | 37.82 s | 2186.8 MB | 103.0 |

### Read Performance (1,000,000 cells)

| Library | Version | Model | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **xlsxrb (Streaming)** | - | Streaming | **3.17 s** | 3.22 s | 91.4 MB | 43.0 |
| **simple_xlsx_reader** | 5.1.0 | Streaming | 4.48 s | 4.45 s | **38.5 MB** | 1669.0 |
| **xlsxrb (In-Memory)** | - | In-Memory | 5.71 s | 5.89 s | 224.8 MB | 63.0 |
| **creek** | 2.6.3 | Streaming | 8.14 s | 8.02 s | 834.6 MB | 477.0 |
| **xsv** | 1.4.1 | Streaming | 14.61 s | 14.50 s | 76.1 MB | 2224.0 |
| **roo** | 3.0.0 | Streaming | 15.69 s | 13.36 s | 119.7 MB | 441.0 |
| **rubyXL** | 3.4.38 | In-Memory | 37.13 s | 40.35 s | 2537.6 MB | 146.0 |

---

## Reproducing Benchmarks Locally

The benchmark suite leverages [`bundler/inline`](https://bundler.io/v2.5/guides/bundler_in_a_single_file_ruby_script.html) to run each library in an isolated subprocess (`Bundler.with_unbundled_env`), eliminating cross-gem pollution and ensuring accurate memory measurements.

To run the suite on your machine:

```bash
ruby benchmark.rb 100000 10
```
