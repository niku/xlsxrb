# Xlsxrb

A Ruby library for reading and writing XLSX files with streaming support.

## Motivation

The Ruby ecosystem already has great XLSX libraries. Each is well-designed for its purpose:

| Library | Read | Write | Streaming (low memory) |
|---------|------|-------|------------------------|
| [roo](https://rubygems.org/gems/roo) | ✅ | ❌ | ✅ |
| [creek](https://rubygems.org/gems/creek) | ✅ | ❌ | ✅ |
| [xsv](https://rubygems.org/gems/xsv) | ✅ | ❌ | ✅ |
| [caxlsx / axlsx](https://rubygems.org/gems/caxlsx) | ❌ | ✅ | ❌ |
| [xlsxtream](https://rubygems.org/gems/xlsxtream) | ❌ | ✅ | ✅ |
| [rubyXL](https://rubygems.org/gems/rubyXL) | ✅ | ✅ | ❌ |
| [fast_excel](https://rubygems.org/gems/fast_excel) | ❌ | ✅ | ✅ |

Each of these libraries makes deliberate tradeoffs, and they do so thoughtfully. Some focus exclusively on highly efficient reading or writing by streaming data, while others provide a rich API for complex, in-memory document modifications.

`xlsxrb` is for cases where you need **both** reading and writing in a single library, while also keeping memory usage predictable for large files.

### Design Principles

- **No Third-Party Runtime Dependencies:** This library avoids third-party XLSX/XML/ZIP gems and relies only on the Ruby standard library and Ruby bundled gems such as `zlib` and `rexml`. This keeps the runtime footprint small while remaining compatible with modern Ruby.
- **Streaming Support:** Both reading and writing are designed to handle large files efficiently by streaming data, keeping memory usage low and predictable.
- **Memory-Efficient XML Parsing:** For reading operations, the library uses REXML's SAX parser instead of DOM-based parsing to avoid loading entire XML documents into memory. This enables true streaming capability for large spreadsheets.
- **Modern Ruby 4.0+:** Built for the future with Ruby 4.0 or higher.

## Installation

```bash
bundle add xlsxrb
```

Or without Bundler:

```bash
gem install xlsxrb
```

On Ruby 4+, some components used by `xlsxrb` and its test suite are shipped as bundled gems rather than built-in default libraries. When using Bundler, those bundled gems are resolved and installed in the usual way.

## Usage

`xlsxrb` offers two different approaches to reading and writing XLSX files: **Streaming** and **In-Memory**.

In most cases, the **Streaming** approach is the best choice because it is highly memory efficient, avoiding loading entire files or structures into RAM. You should always try the Streaming approach first.

However, if your use case requires **Random Access** (e.g., reading a cell at `Z100`, then returning to `A1`) or you need to build or modify an entire document iteratively before writing, the **In-Memory** approach is required.

### 1. Streaming (Recommended)

#### Streaming Read

For large files, use `Xlsxrb.foreach` to read rows one at a time without loading the entire file into memory.

```ruby
require "xlsxrb"

# Yields Xlsxrb::Elements::Row objects for the first sheet
Xlsxrb.foreach("large_file.xlsx", sheet: 0) do |row|
  puts "Row #{row.index}: #{row.cells.map(&:value).join(', ')}"
end
```

#### Streaming Write

To generate large files efficiently, use `Xlsxrb.generate`. This yields a stream writer that writes data directly to the archive.

```ruby
require "xlsxrb"

Xlsxrb.generate("large_output.xlsx") do |writer|
  writer.add_sheet("Sales Data") do
    # Write a header row
    writer.add_row(["Date", "Amount", "Status"])

    # Write data rows
    10_000.times do |i|
      writer.add_row([Date.today - i, i * 100, true])
    end

    # You can also set column widths (0-based index)
    writer.set_column(0, width: 15.5)
  end
end
```

#### Adding Charts (Streaming)

You can generate charts iteratively without loading data into memory by adding them during the streaming process.

```ruby
require "xlsxrb"

Xlsxrb.generate("streaming_chart.xlsx") do |w|
  w.add_sheet("Sales Data") do |s|
    s.add_row(["Month", "Value"])
    s.add_row(["Jan", 100])
    s.add_row(["Feb", 200])

    w.add_chart(
      type: :bar,
      title: "Monthly Sales",
      from_col: 3,
      from_row: 0,
      series: [
        { cat_ref: "Sales Data!$A$2:$A$3", val_ref: "Sales Data!$B$2:$B$3" }
      ]
    )
  end
end
```

### 2. In-Memory

#### Reading an entire file into memory

`Xlsxrb.read` parses the entire XLSX file and returns a workbook object containing all sheets, rows, and cells. Because the entire structure is in memory, you can randomly access any cell by its reference.

```ruby
require "xlsxrb"

workbook = Xlsxrb.read("example.xlsx")
sheet = workbook.sheets.first

# 1. Sequential access
sheet.rows.each do |row|
  puts row.values.join(", ")
end

# 2. Random access by cell reference (Not possible with streaming)
puts "Value at C10: #{sheet.cell_value("C10")}"

# 3. Random access by 0-based row index
row_five = sheet.row_at(4)
puts row_five.cell_at(2).value if row_five
```

#### Writing a workbook from memory

You can create or modify a workbook object and write it out using `Xlsxrb.write`. The `Xlsxrb.build` method provides a convenient DSL to construct the in-memory object hierarchy.

```ruby
require "xlsxrb"

workbook = Xlsxrb.build do |w|
  w.add_sheet("My Sheet") do |s|
    s.add_row(["Hello", "World"])
  end
end

Xlsxrb.write("output.xlsx", workbook)
```

#### Adding Charts (In-Memory)

You can build charts entirely in-memory using the `Xlsxrb.build` DSL, giving you full control over the spreadsheet hierarchy.

```ruby
require "xlsxrb"

workbook = Xlsxrb.build do |w|
  w.add_sheet("Sales Data") do |s|
    s.add_row(["Month", "Value"])
    s.add_row(["Jan", 100])
    s.add_row(["Feb", 200])

    s.add_chart(
      type: :pie,
      title: "Sales Distribution",
      from_col: 3,
      from_row: 0,
      series: [
        { cat_ref: "Sales Data!$A$2:$A$3", val_ref: "Sales Data!$B$2:$B$3" }
      ]
    )
  end
end

Xlsxrb.write("memory_chart.xlsx", workbook)
```

## Styling

You can apply styles to cells in both streaming and in-memory modes. Styles support:
- **Font properties**: bold, italic, size, name, color, underline, strike
- **Fill properties**: pattern (solid, etc.) with colors, or gradients
- **Border properties**: left, right, top, bottom, diagonal with style and color
- **Number format**: custom number formats

### In-Memory Styling

Define styles using the fluent DSL in `Xlsxrb.build`:

```ruby
require "xlsxrb"

workbook = Xlsxrb.build do |w|
  w.add_sheet("Sales") do |s|
    # Define reusable styles
    s.add_style("header") do |style|
      style.bold.size(14).font_color("FFFF0000")  # Red bold, size 14
    end

    s.add_style("total") do |style|
      style.bold.fill_color("FF00FF00")  # Green background, bold
    end

    # Apply styles to rows by specifying style names for each column
    s.add_row(["Date", "Amount", "Status"], styles: ["header", "header", "header"])

    # Add data rows
    s.add_row([Date.today, 1000, "Pending"])
    s.add_row([Date.today - 1, 2000, "Complete"])

    # Apply styles to specific columns in a row
    s.add_row(["Total", 3000, ""], styles: { 0 => "total", 1 => "total" })
  end
end

Xlsxrb.write("styled_output.xlsx", workbook)
```

### Streaming Styling

Define styles in streaming mode with `Xlsxrb.generate`:

```ruby
require "xlsxrb"

Xlsxrb.generate("streaming_styled.xlsx") do |w|
  # Define styles
  w.add_style("header") do |style|
    style.bold.size(12).font_color("FF0000FF")  # Blue bold, size 12
  end

  w.add_style("data") do |style|
    style.fill_color("FFFFC000")  # Orange background
  end

  w.add_sheet("Data") do
    # Apply styles to header row
    w.add_row(["Product", "Qty"],
              styles: { 0 => "header", 1 => "header" })

    # Add data rows with alternating styles
    (1..100).each do |i|
      styles = i % 2 == 0 ? { 0 => "data", 1 => "data" } : nil
      w.add_row(["Item ##{i}", i * 10], styles: styles)
    end
  end
end
```

### StyleBuilder API

The `Xlsxrb::StyleBuilder` class provides a fluent interface for defining styles. Common methods:

**Font methods:**
- `bold(true/false)` — Apply bold formatting
- `italic(true/false)` — Apply italic formatting
- `size(num)` — Set font size (e.g., 12, 14)
- `font_name(name)` — Set font name (e.g., "Arial", "Calibri")
- `font_color(color)` — Set font color (RGB hex, e.g., "FFFF0000" for red)
- `underline(val)` — Set underline style (e.g., "single", "double")
- `strike(true/false)` — Apply strikethrough

**Fill methods:**
- `fill_color(color)` — Solid fill with RGB hex color
- `fill_pattern(pattern, fg_color:, bg_color:)` — Pattern fill
- `fill_gradient(type:, degree:, stops:)` — Gradient fill

**Border methods:**
- `border_all(style:, color:)` — Apply border to all sides
- `border_left(style:, color:)` — Left border only
- `border_right(style:, color:)` — Right border only
- `border_top(style:, color:)` — Top border only
- `border_bottom(style:, color:)` — Bottom border only

**Number format:**
- `number_format(num_fmt_id)` — Apply number format

## Specification

This project aims to be compliant with [ECMA-376](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/) (Office Open XML file formats). Specifically, the library targets the **Transitional** version of the specification rather than the **Strict** version. The Transitional version (detailed in Part 4) is the format most commonly produced and consumed by existing spreadsheet applications, making it the practical choice for real-world interoperability.

## Benchmarks

The following benchmarks measure the time and memory required to process both 100,000 cells (10,000 rows × 10 columns) and 1,000,000 cells (100,000 rows × 10 columns) XLSX files, averaged over 5 iterations on Ruby 3.4+.

### Write Performance (100,000 cells)

| Library | Time | Peak Memory | CPU |
| :--- | ---: | ---: | ---: |
| `xlsxtream` (Streaming) | 0.08 s | 68.0 MB | 90.5 % |
| `fast_excel` (Streaming)* | 0.13 s | 67.1 MB | 95.9 % |
| `caxlsx` (In-Memory) | 0.28 s | 75.6 MB | 84.4 % |
| **`xlsxrb` (Streaming)** | **0.48 s** | **102.8 MB** | **98.6 %** |
| **`xlsxrb` (In-Memory)** | **0.52 s** | **143.6 MB** | **99.7 %** |
| `rubyXL` (In-Memory) | 2.06 s | 273.7 MB | 99.6 % |

*\* `fast_excel` is a C-extension binding to libxlsxwriter, whereas `xlsxrb` is pure Ruby.*

### Read Performance (100,000 cells)

| Library | Time | Peak Memory | CPU |
| :--- | ---: | ---: | ---: |
| `creek` (Streaming) | 0.55 s | 167.8 MB | 99.2 % |
| `roo` (Streaming) | 0.81 s | 88.2 MB | 97.5 % |
| `xsv` (Streaming) | 1.37 s | 95.0 MB | 99.0 % |
| `rubyXL` (In-Memory) | 1.78 s | 282.8 MB | 99.6 % |
| **`xlsxrb` (Streaming)** | **1.97 s** | **72.1 MB** | **99.9 %** |
| **`xlsxrb` (In-Memory)** | **5.12 s** | **145.6 MB** | **99.6 %** |

### Write Performance (1,000,000 cells)

| Library | Time | Peak Memory | CPU |
| :--- | ---: | ---: | ---: |
| `xlsxtream` (Streaming) | 0.14 s | 68.0 MB | 89.9 % |
| `fast_excel` (Streaming)* | 1.17 s | 70.9 MB | 99.1 % |
| `caxlsx` (In-Memory) | 1.62 s | 147.0 MB | 96.9 % |
| **`xlsxrb` (Streaming)** | **3.48 s** | **500.5 MB** | **99.5 %** |
| **`xlsxrb` (In-Memory)** | **5.19 s** | **903.4 MB** | **99.2 %** |
| `rubyXL` (In-Memory) | 39.47 s | 2076.7 MB | 99.1 % |

### Read Performance (1,000,000 cells)

| Library | Time | Peak Memory | CPU |
| :--- | ---: | ---: | ---: |
| `creek` (Streaming) | 5.38 s | 716.6 MB | 98.9 % |
| `roo` (Streaming) | 6.21 s | 132.3 MB | 98.2 % |
| `xsv` (Streaming) | 15.02 s | 101.0 MB | 99.3 % |
| **`xlsxrb` (Streaming)** | **20.43 s** | **114.8 MB** | **99.9 %** |
| `rubyXL` (In-Memory) | 25.05 s | 1849.4 MB | 99.1 % |
| **`xlsxrb` (In-Memory)** | **50.97 s** | **880.8 MB** | **99.5 %** |

*Note: `xlsxrb` is designed for strict OOXML parsing accuracy and full structural mapping, rather than raw read speed. Still, its streaming implementation provides the lowest memory footprint among pure Ruby parsers.*

For reference, the following specification files from the Ecma International website are located in the `vendor/docs/` directory:

- `vendor/docs/ECMA-376-Part1/Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf`: Part 1 - Fundamentals And Markup Language Reference
- `vendor/docs/ECMA-376-Part2/ECMA-376-2_5th_edition_december_2021.pdf`: Part 2 - Open Packaging Conventions
- `vendor/docs/ECMA-376-Part3/ECMA-376-3_5th_edition_december_2015.pdf`: Part 3 - Markup Compatibility and Extensibility
- `vendor/docs/ECMA-376-Part4/Ecma Office Open XML Part 4 - Transitional Migration Features.pdf`: Part 4 - Transitional Migration Features

## Testing Strategy

To ensure high quality and strict compliance with the ECMA-376 specification while maintaining a fast development loop, we employ a tiered testing strategy:

1.  **Unit Tests (Fast & No Dependencies):**
    - Verify individual components (Writer, Reader, Packaging) using only the Ruby standard library and bundled gems.
    - Assert internal state and XML generation logic without performing heavy disk I/O or shelling out to external processes.
    - **Round-Trip Testing:** Ensure that files generated by the `xlsxrb` Writer can be seamlessly and accurately parsed back by the `xlsxrb` Reader. This confirms internal consistency and perfect symmetry between our reading and writing components without external dependencies.

2.  **Interoperability Testing (E2E):**
    - Utilize the official **[Open XML SDK](https://github.com/dotnet/Open-XML-SDK)** for robust, two-way verification:
        - **Writer Validation:** Files generated by `xlsxrb` are read and validated using a script powered by the Open XML SDK to verify the state, correctness of written values, and structural/schema compliance.
        - **Reader Validation:** Complex, real-world XLSX files are programmatically generated by the Open XML SDK, which `xlsxrb` then reads and parses.
    - This approach provides a strong guarantee of structural correctness and compatibility using the standard reference implementation.

**Note on Test Environments:**
While Unit Tests can run anywhere, the Interoperability tests require an external system dependency (`.NET SDK`). All tests are integrated into the provided **Dev Container** environment to ensure a seamless and consistent experience across local machines and CI.

## Development

This project is designed to be developed using [Dev Containers](https://containers.dev/). The provided `.devcontainer` configuration includes all necessary tools (Ruby 4.0, and the .NET SDK for E2E testing) to ensure a consistent development environment.

After checking out the repo, run `bin/setup` to install dependencies. Then, run `rake test` to run the tests. You can also run `bin/console` for an interactive prompt that will allow you to experiment.

### Development Workflow

To ensure systematic progress, perfect round-trip compatibility, and strict adherence to the ECMA-376 specification, we follow this iterative development cycle for each new feature:

1.  **Select a Feature:** Choose a specific element or behavior from the specification to implement.
2.  **Writer Unit Tests:** Write unit tests for the Writer component targeting this feature.
3.  **Writer Implementation:** Implement the Writer functionality.
4.  **Run Writer Tests:** Execute the Writer unit tests. If they fail, return to step 3.
5.  **Writer E2E & Validation:** Test the Writer's generated XLSX file using the Open XML SDK. This includes structural validation using `OpenXmlValidator`. If the test or validation fails, return to step 2.
6.  **Reader Unit Tests:** Write unit tests for the Reader component. **Crucially, include round-trip tests** to ensure the Reader can perfectly parse the output of your Writer.
7.  **Reader Implementation:** Implement the Reader functionality.
8.  **Run Reader Tests:** Execute the Reader unit tests. If they fail, return to step 6 or 7. If the round-trip test reveals a structural flaw in the Writer's output, return all the way back to step 2.
9.  **Reader E2E:** Verify that the Reader can successfully parse a valid XLSX file generated by the Open XML SDK that includes the new feature. If it fails, return to step 6 or 7.
10. **Full Test Suite:** Run the entire test suite (`rake test`). If any tests fail, trace back to the appropriate step.
11. **Commit:** Commit the changes. The commit message must clearly describe the specific feature implemented in this cycle.
12. **Next Feature:** Proceed to the next feature and return to step 1.

#### E2E Policy

E2E tests are required for every new feature. Omitting them is the exception, not the rule, and requires explicit justification.

A strong signal that E2E should not be omitted: if you are adding a new XML element, a new attribute on a top-level structure, or a new public API parameter, E2E is expected.

Omission is only acceptable when **all** of the following hold:

1. The change adds a minor attribute to an XML structure that is **already exercised end-to-end** by an existing E2E scenario for the same element.
2. No new XML element or branch is introduced.
3. Unit tests and round-trip tests fully cover the new behaviour.
4. `rake test` passes with Open XML SDK validation included.
5. The commit message explicitly names the existing E2E scenario that provides coverage and states why a new scenario adds no value.

To install this gem onto your local machine, run `bundle exec rake install`. To release a new version, update the version number in `version.rb`, and then run `bundle exec rake release`, which will create a git tag for the version, push git commits and the created tag, and push the `.gem` file to [rubygems.org](https://rubygems.org).

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/niku/xlsxrb. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).

## License

The gem is available as open source under the terms of the [MIT License](https://opensource.org/licenses/MIT).

## Code of Conduct

Everyone interacting in the Xlsxrb project's codebases, issue trackers, chat rooms and mailing lists is expected to follow the [code of conduct](https://github.com/niku/xlsxrb/blob/main/CODE_OF_CONDUCT.md).
