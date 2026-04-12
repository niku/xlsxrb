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

`add_chart` also supports a block form:

```ruby
Xlsxrb.generate("streaming_chart_block.xlsx") do |w|
  w.add_sheet("Sales") do
    w.add_row(["Month", "Value"])
    w.add_row(["Jan", 100])
    w.add_row(["Feb", 200])

    w.add_chart do |c|
      c.type :bar
      c.title "Monthly Sales"
      c.series(cat_ref: "Sales!$A$2:$A$3", val_ref: "Sales!$B$2:$B$3")
    end
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

`add_style` also supports an options form:

```ruby
s.add_style("header", bold: true, size: 14, font_color: "FFFF0000")
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

## Sheet Features

All sheet-level features work identically in both the streaming (`Xlsxrb.generate`) and in-memory (`Xlsxrb.build`) APIs.
The examples below use `Xlsxrb.generate`; replace it with `Xlsxrb.build` + `Xlsxrb.write` for in-memory use.

### Hyperlinks

Attach a URL or an internal cell reference to a cell.

```ruby
Xlsxrb.generate("links.xlsx") do |w|
  w.add_sheet("Links") do |s|
    s.add_row(["Visit Example", "Jump to cell"])

    # External URL
    s.add_hyperlink("A1", "https://example.com", display: "Example", tooltip: "Open site")

    # Internal location reference (no URL)
    s.add_hyperlink("B1", location: "Sheet2!A1")
  end
end
```

### Auto Filter

Add a filter drop-down to a range of columns.

```ruby
Xlsxrb.generate("filtered.xlsx") do |w|
  w.add_sheet("Data") do |s|
    s.add_row(["Name", "Score"])
    s.add_row(["Alice", 95])
    s.add_row(["Bob", 87])

    s.set_auto_filter("A1:B3")
  end
end
```

### Data Validation

Restrict allowed values in a range of cells.

```ruby
Xlsxrb.generate("validated.xlsx") do |w|
  w.add_sheet("Form") do |s|
    s.add_row(["Rating (1-5)", "Category"])

    # Whole-number range
    s.add_data_validation("A2:A100",
      type: :whole, operator: :between,
      formula1: "1", formula2: "5",
      show_error_message: true,
      error_title: "Invalid", error: "Enter a number from 1 to 5")

    # Drop-down list
    s.add_data_validation("B2:B100",
      type: :list, formula1: '"Alpha,Beta,Gamma"',
      show_error_message: true)
  end
end
```

### Conditional Formatting

Highlight cells automatically based on their value.

```ruby
Xlsxrb.generate("conditional.xlsx") do |w|
  w.add_sheet("Scores") do |s|
    s.add_row([90, 45, 72, 88])

    # Highlight cells greater than 80
    s.add_conditional_format("A1:D1",
      type: :cell_is, operator: :greaterThan,
      formula: "80", priority: 1,
      dxf_id: 0)
  end
end
```

### Tables

Wrap a range in a structured table with column headers and an optional style.

```ruby
Xlsxrb.generate("table.xlsx") do |w|
  w.add_sheet("Report") do |s|
    s.add_row(["Name", "Score"])
    s.add_row(["Alice", 95])
    s.add_row(["Bob", 87])

    s.add_table("A1:B3",
      columns: ["Name", "Score"],
      name: "ScoreTable",
      display_name: "ScoreTable",
      style: "TableStyleMedium9",
      show_first_column: true,
      show_last_column: false)
  end
end
```

### Comments

Attach a text comment (note) to a cell.

```ruby
Xlsxrb.generate("comments.xlsx") do |w|
  w.add_sheet("Notes") do |s|
    s.add_row(["Value"])
    s.add_row([42])

    s.add_comment("A2", "This is the answer", author: "Alice")
  end
end
```

### Merge Cells

Merge a rectangular range of cells into one.

```ruby
Xlsxrb.generate("merged.xlsx") do |w|
  w.add_sheet("Layout") do |s|
    s.add_row(["Title", nil, nil])
    s.add_row([1, 2, 3])

    s.merge_cells("A1:C1")
  end
end
```

### Freeze / Split Panes

Lock rows or columns in place while scrolling.

```ruby
Xlsxrb.generate("frozen.xlsx") do |w|
  w.add_sheet("Big Table") do |s|
    s.add_row(["ID", "Name", "Score"])
    100.times { |i| s.add_row([i + 1, "Row #{i + 1}", i * 10]) }

    # Freeze the first row (row index 1 = keep row 0 visible)
    s.set_freeze_pane(row: 1, col: 0)

    # Or split at pixel offsets (non-frozen)
    # s.set_split_pane(x_split: 2000, y_split: 600, top_left_cell: "B5")
  end
end
```

Set the active cell selection (optional, often used together with a freeze pane):

```ruby
s.set_selection("B2", sqref: "B2", pane: "bottomRight")
```

### Page Margins, Page Setup, and Header/Footer

Control how the sheet looks when printed.

```ruby
Xlsxrb.generate("print_ready.xlsx") do |w|
  w.add_sheet("Report") do |s|
    s.add_row(["Header", "Data"])

    # Page margins in inches
    s.set_page_margins(left: 0.75, right: 0.75, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5)

    # Page setup (orientation, paper size, scaling)
    s.set_page_setup(orientation: :landscape, paper_size: 9, scale: 90, fit_to_page: true)

    # Header and footer text (uses Excel format codes)
    s.set_header_footer(
      odd_header: "&L&\"Arial,Bold\"My Company&C&18Report Title&RPage &P of &N",
      odd_footer: "&LConfidential&C&D&RGenerated by xlsxrb"
    )
  end
end
```

### Print Options

Enable or disable specific print settings.

```ruby
Xlsxrb.generate("print_opts.xlsx") do |w|
  w.add_sheet("Grid") do |s|
    s.add_row(["A", "B"])

    s.set_print_option(:grid_lines, true)
    s.set_print_option(:horizontal_centered, true)
    s.set_print_option(:vertical_centered, false)
  end
end
```

### Sheet Protection

Protect a sheet to prevent accidental edits (optionally with a password).

```ruby
Xlsxrb.generate("protected.xlsx") do |w|
  w.add_sheet("Locked") do |s|
    s.add_row(["Read-only data"])

    s.set_sheet_protection(sheet: true, objects: true, scenarios: true)

    # With a password hash (SHA-512 or legacy XOR hash supported by Excel)
    # s.set_sheet_protection(sheet: true, password: "secret")
  end
end
```

### Row and Column Page Breaks

Insert manual page breaks before specific rows or columns.

```ruby
Xlsxrb.generate("breaks.xlsx") do |w|
  w.add_sheet("Pages") do |s|
    20.times { |i| s.add_row(["Row #{i + 1}"]) }

    s.add_row_break(10)  # page break before row 10 (0-based)
    s.add_col_break(3)   # page break before column 3 (0-based, = column D)
  end
end
```

### Images

Embed an image anchored to a cell range.

```ruby
png_bytes = File.binread("logo.png")

Xlsxrb.generate("with_image.xlsx") do |w|
  w.add_sheet("Cover") do |s|
    s.add_row(["Product Report"])

    s.add_image(png_bytes, ext: "png",
      from_col: 0, from_row: 1,
      to_col: 4,  to_row: 10)
  end
end
```

### Shapes (VML Drawing)

Add a basic shape (rectangle, ellipse, etc.) to a sheet.

```ruby
Xlsxrb.generate("shapes.xlsx") do |w|
  w.add_sheet("Diagram") do |s|
    s.add_row(["See the shape below"])

    s.add_shape(
      preset: "rect",
      text: "Important!",
      from_col: 0, from_row: 2,
      to_col: 3,   to_row: 6,
      name: "Banner",
      fill_color: "#FFFFC0",
      line_color: "#FF0000"
    )
  end
end
```

### Sheet Properties and View Settings

Set per-sheet visual properties (tab colour, grid lines, zoom level, etc.).

```ruby
Xlsxrb.generate("styled_sheet.xlsx") do |w|
  w.add_sheet("Summary") do |s|
    s.add_row(["Data"])

    s.set_sheet_property(:tab_color, "FF4472C4")   # blue tab
    s.set_sheet_view(:show_grid_lines, false)
    s.set_sheet_view(:zoom_scale, 120)
  end
end
```

## Workbook Features

Workbook-level methods are called directly on the writer/builder object (not inside `add_sheet`).

### Defined Names, Print Area, and Print Titles

```ruby
Xlsxrb.generate("named.xlsx") do |w|
  w.add_sheet("Data") do |s|
    s.add_row(["Name", "Score"])
    s.add_row(["Alice", 95])
  end

  # General named range
  w.add_defined_name("TaxRate", "0.1", hidden: false)

  # Print area for the sheet named "Data"
  w.set_print_area("A1:B10", sheet: "Data")

  # Repeat first row and first column when printing
  w.set_print_titles(rows: "1:1", cols: "A:A", sheet: "Data")
end
```

### Workbook Protection

Lock the workbook structure (prevents adding/removing/renaming sheets).

```ruby
Xlsxrb.generate("locked_workbook.xlsx") do |w|
  w.add_sheet("Sheet1") { |s| s.add_row(["Data"]) }

  w.set_workbook_protection(lock_structure: true, lock_windows: false)
end
```

### Document Properties (Core, App, and Custom)

Embed metadata into the XLSX file's document properties panels.

```ruby
Xlsxrb.generate("with_props.xlsx") do |w|
  w.add_sheet("Sheet1") { |s| s.add_row(["Hello"]) }

  # Core properties (Dublin Core — visible in File > Info)
  w.set_core_property(:title,    "Quarterly Report")
  w.set_core_property(:subject,  "Sales data Q1 2026")
  w.set_core_property(:creator,  "Alice")
  w.set_core_property(:keywords, "sales, report, 2026")
  w.set_core_property(:description, "Auto-generated by xlsxrb")

  # App properties (application-level metadata)
  w.set_app_property(:application, "MyApp/2.0")
  w.set_app_property(:company,     "Acme Corp")

  # Custom properties (arbitrary key/value pairs)
  w.add_custom_property("ReportVersion", "42", type: :integer)
  w.add_custom_property("ApprovedBy",    "Bob", type: :string)
  w.add_custom_property("Published",     true,  type: :bool)
end
```

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

High-level API expansion follows the Facade rules documented in [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md). In short: if a low-level writer feature is stable, the default expectation is that it should eventually be exposed through the high-level DSL as well, with consistent naming, both streaming and in-memory coverage, backward-compatible options/block forms where practical, and matching Facade-level tests.

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
