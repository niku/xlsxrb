# Xlsxrb

A Ruby library for reading and writing XLSX files with streaming support.

## Motivation

The Ruby ecosystem already has great XLSX libraries. Each is well-designed for its purpose:

| Library                                            | Read | Write | Streaming (low memory) |
| -------------------------------------------------- | ---- | ----- | ---------------------- |
| [roo](https://rubygems.org/gems/roo)               | ✅   | ❌    | ✅                     |
| [creek](https://rubygems.org/gems/creek)           | ✅   | ❌    | ✅                     |
| [xsv](https://rubygems.org/gems/xsv)               | ✅   | ❌    | ✅                     |
| [caxlsx / axlsx](https://rubygems.org/gems/caxlsx) | ❌   | ✅    | ❌                     |
| [xlsxtream](https://rubygems.org/gems/xlsxtream)   | ❌   | ✅    | ✅                     |
| [rubyXL](https://rubygems.org/gems/rubyXL)         | ✅   | ✅    | ❌                     |
| [fast_excel](https://rubygems.org/gems/fast_excel) | ❌   | ✅    | ✅                     |

Each of these libraries makes deliberate tradeoffs, and they do so thoughtfully. Some focus exclusively on highly efficient reading or writing by streaming data, while others provide a rich API for complex, in-memory document modifications.

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
Generate large files efficiently by writing data directly to the file stream:
```ruby
require "xlsxrb"

Xlsxrb.generate("large_output.xlsx") do |writer|
  writer.add_sheet("Sales Data") do
    writer.add_row(["Date", "Amount", "Status"])
    writer.add_row([Date.today, 100, true])
    writer.set_column(0, width: 15.5)
  end
end
```

#### Streaming Read
Read rows one at a time without loading the entire file into memory:
```ruby
require "xlsxrb"

Xlsxrb.foreach("large_file.xlsx", sheet: 0) do |row|
  puts "Row #{row.index}: #{row.cells.map(&:value).join(', ')}"
end
```

*(For In-Memory document building, cell modifications, or template updating, please refer to the detailed RDoc API documentation).*

## Feature Support & ECMA-376 Compliance

`xlsxrb` is designed for full interoperability and strict compliance with the ECMA-376 (Office Open XML) Transitional specification. It supports nearly all major spreadsheet features required for business reports:

* Cells & Layout: Formulas, Hyperlinks, Merge Cells, Freeze & Split Panes, Page Setup (margins, headers/footers, scaling, gridlines).
* Data & Controls: Auto Filters, Data Validations (dropdowns, range limits), Sheet Protection.
* Formatting & Styling: Rich Text, Cell Tables, Conditional Formatting (color scales, data bars, icon sets).
* Graphics & Charts: Embedded Images, Shapes & Drawings, Sparklines, Charts (Line, Bar, Pie, Area, Radar, Scatter).
* Workbook Level: Defined Names, Print Areas, Workbook Protection, and Document Metadata (core, app, custom properties).

For detailed specification references and policies, see [SPEC_SOURCES.md](docs/SPEC_SOURCES.md).

## Benchmarks

The following benchmarks measure the time and memory required to process a 1,000,000 cells (100,000 rows × 10 columns) spreadsheet, demonstrating `xlsxrb`'s memory consumption and processing times.

### Write Performance (1,000,000 cells)

| Library                | Time    | Peak Memory | GC Count |
| ---------------------- | ------- | ----------- | -------- |
| xlsxtream (Streaming)  | 0.12 s  | 64.9 MB     | 9.0      |
| fast_excel (Streaming) | 1.34 s  | 64.5 MB     | 28.0     |
| caxlsx (In-Memory)     | 2.64 s  | 142.7 MB    | 16.0     |
| xlsxrb (Streaming)     | 3.32 s  | 216.9 MB    | 65.0     |
| xlsxrb (In-Memory)     | 6.23 s  | 431.2 MB    | 68.0     |
| rubyXL (In-Memory)     | 37.16 s | 2105.9 MB   | 90.0     |

### Read Performance (1,000,000 cells)

| Library            | Time    | Peak Memory | GC Count |
| ------------------ | ------- | ----------- | -------- |
| xlsxrb (Streaming) | 5.38 s  | 101.8 MB    | 1429.0   |
| creek (Streaming)  | 7.05 s  | 706.2 MB    | 3985.0   |
| roo (Streaming)    | 7.24 s  | 128.1 MB    | 279.0    |
| xlsxrb (In-Memory) | 8.62 s  | 996.1 MB    | 28.0     |
| xsv (Streaming)    | 14.41 s | 93.5 MB     | 998.0    |
| rubyXL (In-Memory) | 24.58 s | 1856.8 MB   | 127.0    |

### Running the Benchmarks Locally

The benchmark data is gathered using the bundled [benchmark.rb](file:///workspaces/xlsxrb/benchmark.rb) script, which runs each library's code in isolated subprocesses to ensure accurate memory and GC measurements.

To run the benchmark locally for 1,000,000 cells (100,000 rows):
```bash
ruby benchmark.rb 100000
```

## Testing & Quality Assurance

To support reliability, compliance with the ECMA-376 specification, and consistent updates, `xlsxrb` is backed by a 4-tier testing strategy:

1. Round-Trip Testing: Unit tests verify that every generated sheet can be reliably parsed back by the reader with identical content and styling.
2. Contract Consistency: The library ensures semantic output consistency between the Streaming (`Xlsxrb.generate`) and In-Memory (`Xlsxrb.build`) APIs.
3. Official Open XML SDK Validation (E2E): Every feature is validated against the official Microsoft Open XML SDK. Generated spreadsheets are structurally checked to prevent file corruption warnings.
4. Visual Regression Testing (VRT): To guarantee rendering correctness, generated XLSX files are rendered via a headless LibreOffice Calc engine, and compared pixel-by-pixel against visual baselines.

For details on running the tests locally or within our pre-configured Dev Container, see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

## Development

We welcome contributions! The project is configured with a ready-to-use Dev Container to streamline local environment setup.

For contribution guidelines, E2E testing policies, and the step-by-step development workflow, please refer to [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/niku/xlsxrb. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Xlsxrb project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).
