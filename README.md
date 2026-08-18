# Xlsxrb

A Ruby library for reading and writing XLSX files with streaming support.

## Motivation

The Ruby ecosystem already has great XLSX libraries. Each is well-designed for its purpose:

| Library                                                    | Read | Write | Model                 | Write String Storage | Rich Formatting |
| ---------------------------------------------------------- | ---- | ----- | --------------------- | -------------------- | --------------- |
| [roo](https://rubygems.org/gems/roo)                       | ✅   | ❌    | Streaming             | N/A (Read-only)      | ⚠️ (Formulas, Basic styles) |
| [creek](https://rubygems.org/gems/creek)                   | ✅   | ❌    | Streaming             | N/A (Read-only)      | ❌ (Raw cell values) |
| [xsv](https://rubygems.org/gems/xsv)                       | ✅   | ❌    | Streaming             | N/A (Read-only)      | ❌ (Fast plain text) |
| [simple_xlsx_reader](https://rubygems.org/gems/simple_xlsx_reader) | ✅ | ❌ | Streaming         | N/A (Read-only)      | ❌ (Plain data & types) |
| [caxlsx / axlsx](https://rubygems.org/gems/caxlsx)         | ❌   | ✅    | In-Memory             | Inline (opt: SST)    | ✅ (Charts, Styles) |
| [write_xlsx](https://rubygems.org/gems/write_xlsx)         | ❌   | ✅    | In-Memory             | SST                  | ✅ (Charts, Styles) |
| [xlsxtream](https://rubygems.org/gems/xlsxtream)           | ❌   | ✅    | Streaming             | Inline (opt: SST)    | ❌ (Plain data only) |
| [fast_excel](https://rubygems.org/gems/fast_excel)         | ❌   | ✅    | Streaming (C Ext)     | SST (opt: Inline)    | ⚠️ (Basic styles) |
| [rubyXL](https://rubygems.org/gems/rubyXL)                 | ✅   | ✅    | In-Memory             | Inline / Direct      | ✅ (DOM editing) |
| **[xlsxrb](https://github.com/niku/xlsxrb)**               | ✅   | ✅    | **Streaming / In-Memory** | **SST**          | ✅ **(Full Features)** |

Each of these libraries makes deliberate tradeoffs, and they do so thoughtfully:
* **Memory & Execution Model (Streaming vs In-Memory)**: Streaming libraries write or read rows sequentially on-the-fly to maintain a constant, low-memory footprint regardless of row count. In-memory libraries build complete document object trees, offering flexible random access and cell updates at the cost of high RAM usage on large sheets.
* **String Storage Architecture (SST vs Inline Strings)**:
  * **SST (Shared String Table)**: De-duplicates strings into a central dictionary (`xl/sharedStrings.xml`), referencing them by numeric IDs in cell entries (`<c t="s"><v>0</v></c>`). This is standard Microsoft Excel behavior, producing significantly smaller raw XML documents (50–100% smaller) and reducing Excel's memory footprint when opening spreadsheets.
  * **Inline Strings**: Writes text directly into cell payloads (`<c t="inlineStr"><is><t>...</t></is></c>`). Bypassing the dictionary enables blazing-fast raw throughput for simple data exports, but inflates uncompressed XML size and limits advanced formatting (e.g. styling, cell merges, charts).

Traditionally, attempting to build a "complete package" that offers both reading and writing, rich features, high performance, strict compatibility, and comprehensive documentation presents an inherent open-source challenge: the cumulative maintenance overhead often exceeds the capacity of individual human maintainers.

`xlsxrb` is born from a different premise. We believe that Advanced Agentic AI (AI Coders) can help manage this maintenance demand. By utilizing AI agents to automate rigorous E2E testing, visual regression testing, specification compliance checks, and documentation updates, we can reconcile these competing engineering requirements. This allows us to build and continuously maintain a feature-rich, high-performance, and deeply compatible "all-in-one" XLSX library that remains sustainable for the long run.

### Design Principles

- Minimal Dependencies (Zero Core Logic Dependencies): This library avoids heavy third-party XLSX/XML/ZIP gems, building all core parsing and writing features purely on the Ruby standard library and bundled gems (`zlib`, `rexml`, etc.). The only runtime dependency is `opentelemetry-api`, which provides zero-overhead observability. If you do not configure an OpenTelemetry SDK in your application, it acts as a lightweight no-op, keeping the runtime footprint extremely small.
- Streaming Support: Both reading and writing are designed to handle large files efficiently by streaming data, keeping memory usage low and predictable.
- Memory-Efficient XML Parsing: For reading operations, the library uses a custom byte-level streaming parser for worksheet rows (with targeted SAX parsing where appropriate) instead of DOM-based parsing, so entire XML documents are never loaded into memory. This enables true streaming capability for large spreadsheets.
- Strict Microsoft Excel & OpenXML Interoperability: It is designed to closely follow the Microsoft Office implementation of the ISO 29500 standard. We ensure absolute bidirectional compatibility (both reading and writing) with Microsoft Excel by continuously validating files against the official Microsoft [Open XML SDK](https://github.com/dotnet/Open-XML-SDK).
- AI-Agent Assisted Maintenance (Managing the Engineering Tradeoff): Building a library that is specification-compliant, rich in features, highly compatible, well-documented, and extremely fast typically presents a substantial maintenance challenge. `xlsxrb` addresses this inherent constraint by leveraging Advanced Agentic AI (AI Coders) to automate testing, feature expansion, and compatibility verification. This AI-assisted development process supports the project's long-term sustainability and high software quality.
- Modern Ruby 4.0+: Built for the future with Ruby 4.0 or higher.

## Installation

```bash
bundle add xlsxrb
```

Or without Bundler:

```bash
gem install xlsxrb
```

On Ruby 4+, some components used by `xlsxrb` and its test suite are shipped as bundled gems rather than built-in default libraries. When using Bundler, those bundled gems are resolved and installed in the usual way.

## Interactive Playground (WebAssembly)

You can try `xlsxrb` directly in your browser without installing anything!

[👉 Try the Live Demo / Interactive Playground](https://niku.github.io/xlsxrb/docs/visual/VisualGallery_md.html)

We have integrated an interactive WebAssembly-powered playground into our RDoc documentation. You can edit the code examples, run them in the browser sandbox, and download the generated `.xlsx` spreadsheets immediately.

To launch the playground locally:
1. Generate the WebAssembly bundle and interactive RDoc:
   ```bash
   bundle exec rake doc
   ```
2. Start the local preview server:
   ```bash
   bundle exec rake doc:preview
   ```
3. Open [http://localhost:8000](http://localhost:8000) in your browser, hover over any code block, and click the "Live Preview" or "Download XLSX" buttons!

## Usage

`xlsxrb` supports both low-memory Streaming (recommended for large files) and full In-Memory document manipulation (for random-access cell modifications or updating existing sheets).

For visual demonstrations of various features, check the [Visual Examples Gallery](docs/visual/VisualGallery.md).

### Quick Start: Streaming (Recommended)

#### Streaming Write
Generate large files efficiently with $O(1)$ constant memory by writing data directly to the stream:
```ruby
require "xlsxrb"

Xlsxrb.write("large_output.xlsx") do |writer|
  writer.sheet("Sales Data") do |sheet|
    sheet.row(["Date", "Amount", "Status"])
    sheet.row([Date.today, 100, true])
    sheet.column(0, width: 15.5)
  end
end
```

#### Streaming Read
Read rows and cells lazily one at a time with $O(1)$ constant memory (even for wide sheets with thousands of columns):
```ruby
require "xlsxrb"

# Stream row-by-row and cell-by-cell (O(1) memory)
Xlsxrb.read("large_file.xlsx") do |sheet|
  sheet.each_row do |row|
    row.each_cell do |cell|
      puts "#{cell.ref}: #{cell.value}"
    end
  end
end

# Or stream all cells across the sheet directly
Xlsxrb.read("large_file.xlsx") do |sheet|
  sheet.each_cell do |cell|
    puts "#{cell.ref} = #{cell.value}"
  end
end
```

### Ruby-Idiomatic Core APIs

`xlsxrb` provides clean, standard Ruby interfaces (`Enumerable`, `Row#to_a`, `sheet["A1"]`) that feel natural to every Ruby developer without learning complex library-specific APIs:

#### Reading Spreadsheets
```ruby
require "xlsxrb"

# 1. Read from file path, IO, or raw binary string (O(1) constant memory streaming)
workbook = Xlsxrb.read("data.xlsx")
sheet = workbook.sheets.first

# 2. Extract sheet data into 2D array of values via standard Enumerable
matrix = sheet.map(&:to_a) # => [["Name", "Score"], ["Alice", 100], ["Bob", 95]]

# 3. Explicitly load into memory for coordinate random access (e.g. sheet["A1"])
doc_sheet = sheet.load
doc_sheet["A1"]         # => #<Xlsxrb::Elements::Cell value="Name" ...>
doc_sheet["A1"].value   # => "Name"
```

#### Writing & In-Memory Export (Rails & Mailers)
```ruby
# Build workbook
wb = Xlsxrb.build do |b|
  b.sheet("Report") do |s|
    s.row(["Metric", "Value"])
    s.row(["Users", 1000])
  end
end

# Save directly to file:
Xlsxrb.write("report.xlsx", wb)

# Or export to binary string (ideal for Rails send_data & ActionMailer):
binary_data = Xlsxrb.write(wb)
```

### In-Memory Building & Modifying

`xlsxrb` provides a powerful, immutable-by-default API for modifying existing Excel files or building templates in-memory. 

#### Modifying an Existing File
You can update specific cells or sheets using the functional `Xlsxrb.modify` API, which yields the parsed `Elements::Workbook`.

```ruby
require "xlsxrb"

# Create a template.xlsx for this example
Xlsxrb.build { |builder| builder.sheet("Invoice") }.write("template.xlsx")

Xlsxrb.modify("template.xlsx", "output.xlsx") do |workbook|
  workbook.update_sheet("Invoice") do |sheet|
    # Update specific cells (returns updated sheet)
    sheet = sheet.update_cell("C4", value: "INV-10042")
    sheet = sheet.update_cell("C5", value: Date.today)
    
    # Or append new rows
    sheet.with(rows: sheet.rows + [
      Xlsxrb::Elements::Row.new(index: sheet.rows.size, cells: [])
    ])
  end
end
```

#### Hash & Range Styling (Syntactic Sugar)
You can directly apply inline styles or use Ranges for multiple columns without boilerplate:

```ruby
Xlsxrb.build do |builder|
  # Use [] accessor for sheets
  builder["Report"].row(
    ["ID", "Name", "Score", "Rank"],
    # Apply 'header' style to first two columns, and bold inline style to the third
    styles: { 0..1 => "header", 2 => { font: { bold: true, color: "red" } } }
  )
  
  # Set multiple column widths at once using Ranges
  builder["Report"].column("A".."D", width: 15.0)
end
```

### IDE Autocompletion & Ruby LSP Support

`xlsxrb` bundles a native **Ruby LSP Add-on** (`RubyLsp::Xlsxrb::Addon`) and full **RBS signatures**, enabling zero-configuration method autocompletion and rich Markdown documentation in VS Code and other LSP-enabled editors.

Whether you use standard descriptive block variable names (`|stream_writer|`, `|sheet|`, `|workbook|`) or short names (`|wb|`, `|s|`), your editor will automatically provide complete method suggestions and parameter hints:

```ruby
Xlsxrb.write("output.xlsx") do |writer| # or |wb|
  writer.sheet("Data") do |sheet|      # or |s|
    sheet.row(["Product", "Price"], styles: :bold)
    sheet.auto_filter("A1:B100")
  end
end
```

## Feature Support & ECMA-376 Compliance

`xlsxrb` is designed for full interoperability and strict compliance with the ECMA-376 (Office Open XML) Transitional specification. It supports nearly all major spreadsheet features required for business reports:

* Cells & Layout: Formulas, Hyperlinks, Merge Cells, Freeze & Split Panes, Page Setup (margins, headers/footers, scaling, gridlines).
* Data & Controls: Auto Filters, Data Validations (dropdowns, range limits), Sheet Protection.
* Formatting & Styling: Rich Text, Cell Tables, Conditional Formatting (color scales, data bars, icon sets).
* Graphics & Charts: Embedded Images, Shapes & Drawings, Sparklines, Charts (Line, Bar, Pie, Area, Radar, Scatter).
* Workbook Level: Defined Names, Print Areas, Workbook Protection, and Document Metadata (core, app, custom properties).

For detailed specification references and policies, see [SPEC_SOURCES.md](docs/SPEC_SOURCES.md).

## Benchmarks

The following benchmarks measure the time, peak memory, and GC count required to process a 1,000,000 cells (100,000 rows × 10 columns) spreadsheet across popular Ruby Excel libraries. Each test is executed across 3 independent runs in isolated subprocesses; median values are reported along with the mean execution time.

### Write Performance (1,000,000 cells)

| Library                | Model       | Write String Storage | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| ---------------------- | ----------- | -------------------- | ------------- | ----------- | ----------- | -------- |
| xlsxtream 3.1.0        | Streaming   | Inline String        | 1.19 s        | 1.20 s      | 18.2 MB     | 1072.0   |
| xlsxrb (Streaming)     | Streaming   | SST (Shared)         | 1.73 s        | 1.65 s      | 94.4 MB     | 39.0     |
| fast_excel 0.5.0 (C)   | Streaming   | SST (Shared)         | 1.89 s        | 1.89 s      | 148.2 MB    | 245.0    |
| xlsxrb (In-Memory)     | In-Memory   | SST (Shared)         | 3.84 s        | 3.83 s      | 278.3 MB    | 32.0     |
| write_xlsx 1.15.0      | In-Memory   | SST (Shared)         | 4.32 s        | 4.34 s      | 201.2 MB    | 33.0     |
| caxlsx 4.5.0           | In-Memory   | Inline String        | 5.15 s        | 5.12 s      | 188.6 MB    | 23.0     |
| rubyXL 3.4.38          | In-Memory   | Inline String        | 38.81 s       | 37.82 s     | 2186.8 MB   | 103.0    |

> **Note**: All libraries are evaluated in their **default, out-of-the-box configuration**. Under the same Microsoft Excel-standard Shared String Table (SST) architecture, Pure Ruby `xlsxrb` (Streaming: 1.73s, In-Memory: 3.84s) writes 1,000,000 cells faster than the C-extension `fast_excel` (1.89s) and in-memory gems like `write_xlsx` (4.32s) and `caxlsx` (5.15s).

### Read Performance (1,000,000 cells)

| Library                  | Model       | Time (Median) | Time (Mean) | Peak Memory | GC Count |
| ------------------------ | ----------- | ------------- | ----------- | ----------- | -------- |
| xlsxrb (Streaming)       | Streaming   | 3.17 s        | 3.22 s      | 91.4 MB     | 43.0     |
| simple_xlsx_reader 5.1.0 | Streaming   | 4.48 s        | 4.45 s      | 38.5 MB     | 1669.0   |
| xlsxrb (In-Memory)       | In-Memory   | 5.71 s        | 5.89 s      | 224.8 MB    | 63.0     |
| creek 2.6.3              | Streaming   | 8.14 s        | 8.02 s      | 834.6 MB    | 477.0    |
| xsv 1.4.1                | Streaming   | 14.61 s       | 14.50 s     | 76.1 MB     | 2224.0   |
| roo 3.0.0                | Streaming   | 15.69 s       | 13.36 s     | 119.7 MB    | 441.0    |
| rubyXL 3.4.38            | In-Memory   | 37.13 s       | 40.35 s     | 2537.6 MB   | 146.0    |

### Running the Benchmarks Locally (Reproducibility)

The benchmark suite leverages [`bundler/inline`](https://bundler.io/v2.5/guides/bundler_in_a_single_file_ruby_script.html) to automatically manage and download all peer ecosystem gems without modifying the project's core `Gemfile` or requiring manual global `gem install` steps. Each library is executed in an isolated subprocess (`Bundler.with_unbundled_env`) across multiple runs with standard business dataset rows (integers, strings, floats, booleans, dates) to ensure clean memory and GC measurements without cross-contamination.

To run the complete benchmark suite:
```bash
ruby benchmark.rb 100000 10
```

## Security (Protection against CSV/Excel Injection)

Unlike CSV files which lack type definitions and force Excel to guess types (often inadvertently executing strings starting with `=`), `.xlsx` files generated by `xlsxrb` are strictly typed.

When you pass a Ruby `String` to `xlsxrb`, it explicitly writes it as a `String` (`t="s"`) into the OOXML file. Therefore, **even if a string starts with `=`, Excel will never evaluate it as a formula**. To write a formula, you must explicitly use `Xlsxrb::Elements::Formula.new`. This design completely mitigates CSV/Formula Injection vulnerabilities by default without requiring additional sanitization.

### External Link Updates (`update_links`)

As an extra layer of "defense in depth", `xlsxrb` configures the workbook to **never automatically update external links** when opened (`updateLinks="never"`). This is intentionally set to `never` by default to prevent Excel from silently reaching out to external resources or executing DDE (Dynamic Data Exchange) links, which is a known vector for malware.

If you absolutely need external links to update automatically, you can explicitly override this (though **it is highly discouraged due to security risks**):

```ruby
Xlsxrb.write("file.xlsx") do |wb|
  # WARNING: Enabling this can expose users to malicious external reference vulnerabilities!
  wb.workbook_property(:update_links, "always") 
  # ...
end
```

## Testing & Quality Assurance

To support reliability, compliance with the ECMA-376 specification, and consistent updates, `xlsxrb` is backed by a highly rigorous, enterprise-grade Quality Assurance (QA) and testing architecture.

### Multi-Tier Testing Strategy
* **Round-Trip Testing**: Unit tests verify that every generated sheet can be reliably parsed back by the reader with identical content and styling.
* **Contract Consistency**: Ensures semantic output consistency between the Streaming (`Xlsxrb.write`) and In-Memory (`Xlsxrb.build`) APIs.
* **Property-Based Testing (PBT)**: Automatically generates random data to catch edge cases (e.g., huge numbers, special characters) preventing unexpected crashes.
* **Concurrency Validation**: Thread and Ractor safety checks to guarantee no global variable pollution during parallel execution.
* **Security & DoS Protection**: Hardened against malicious files, including memory exhaustion (ZIP Bombs) and infinite parsing loops.

### Strict Interoperability & Rendering
* **Official Open XML SDK Validation (E2E)**: Every generated spreadsheet is structurally validated against the official Microsoft Open XML SDK to prevent file corruption warnings in Microsoft Excel.
* **Visual Regression Testing (VRT)**: Spreadsheets are rendered via a headless LibreOffice Calc engine and compared pixel-by-pixel against visual baselines to catch subtle rendering regressions.

### Performance & Types
* **Continuous Benchmarking**: Memory usage and processing speeds are profiled in CI on large datasets to prevent performance regressions and OOM leaks.
* **Runtime Type Validation**: Strong dynamic typing using `RBS::Test` to ensure the library's types are perfectly sound at runtime.

For a comprehensive breakdown of our QA matrix, see [docs/QUALITY_ASSURANCE.md](docs/QUALITY_ASSURANCE.md). For details on running tests locally, see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Development

We welcome contributions! The project is configured with a ready-to-use Dev Container to streamline local environment setup.

For contribution guidelines, E2E testing policies, and the step-by-step development workflow (including how to run the Dev Container from your terminal), please refer to [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/niku/xlsxrb. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Xlsxrb project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](CODE_OF_CONDUCT.md).
