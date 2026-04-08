# Xlsxrb Architecture

Xlsxrb uses a multi-layered architecture to separate low-level OpenXML specification details from a high-level, idiomatic Ruby API. This separation of concerns ensures the library is both robust against the complex OpenXML spec and user-friendly for Ruby developers.

## Dependency Constraints

Xlsxrb uses **only** the Ruby standard library and Bundled Gems:

| Library | Purpose |
| :--- | :--- |
| `rexml` (bundled gem) | SAX2-based XML parsing and DOM-based XML generation |
| `zlib` (stdlib) | ZIP deflate/inflate compression |
| `stringio` (stdlib) | In-memory IO for streaming |
| `date` (stdlib) | Excel serial-number ↔ `Date` / `Time` conversion |
| `openssl`, `securerandom` (stdlib) | Password hashing for sheet/workbook protection |

**No third-party gems** (e.g. Nokogiri, rubyzip) are permitted as runtime dependencies.

---

## Directory Structure

```
lib/
  xlsxrb.rb                          # Facade: Xlsxrb.read / .write / .foreach / .generate / .build
  xlsxrb/
    version.rb                       # Xlsxrb::VERSION
    elements.rb                      # Requires for Elements layer
    ooxml.rb                         # Requires for Ooxml layer
    ooxml/                           # Layer 1 – Low-level OOXML
      reader.rb                      #   Xlsxrb::Ooxml::Reader (core reading logic)
      writer.rb                      #   Xlsxrb::Ooxml::Writer (core writing logic)
      zip_generator.rb               #   Xlsxrb::Ooxml::ZipGenerator
      utils.rb                       #   Xlsxrb::Ooxml::Utils (date/time/hash helpers)
      zip_reader.rb                  #   Xlsxrb::Ooxml::ZipReader
      zip_writer.rb                  #   Xlsxrb::Ooxml::ZipWriter
      xml_parser.rb                  #   Xlsxrb::Ooxml::XmlParser
      xml_builder.rb                 #   Xlsxrb::Ooxml::XmlBuilder
      shared_strings_parser.rb       #   Streaming SST reader
      styles_parser.rb               #   styles.xml reader
      worksheet_parser.rb            #   Streaming sheetN.xml reader
      worksheet_writer.rb            #   Streaming sheetN.xml writer
      workbook_parser.rb             #   workbook.xml reader
      workbook_writer.rb             #   workbook.xml / styles / SST writer
    elements/                        # Layer 2 – Domain model
      types.rb                       #   Xlsxrb::Elements::Formula, CellError, RichText
      cell.rb                        #   Xlsxrb::Elements::Cell
      row.rb                         #   Xlsxrb::Elements::Row
      column.rb                      #   Xlsxrb::Elements::Column
      worksheet.rb                   #   Xlsxrb::Elements::Worksheet
      workbook.rb                    #   Xlsxrb::Elements::Workbook
```

---

## Core Architecture

The library is structured into three distinct layers:

### 1. Low-Level Infrastructure (The "OOXML" Layer)

**Namespace:** `Xlsxrb::Ooxml`

**Responsibility:**
This layer directly handles ZIP extraction, XML parsing (via SAX), and XML generation. It adheres strictly to the ECMA-376 OpenXML specification.

* **`Xlsxrb::Ooxml::ZipReader`**: Reads a `.xlsx` ZIP archive entry-by-entry. Accepts a file path or `IO` object. Yields `(entry_name, io)` pairs without loading the entire archive into memory.
* **`Xlsxrb::Ooxml::ZipWriter`**: Streams ZIP local-file-headers and a central directory to a file path or `IO`. Each entry is compressed with `Zlib::Deflate` in a single pass.
* **`Xlsxrb::Ooxml::XmlParser`**: Thin wrapper around `REXML::Parsers::SAX2Parser`. Converts SAX2 events into a Hash/Array tree. Unknown elements are collected as opaque `{ tag:, attrs:, children: }` hashes (see *unmapped_data* below).
* **`Xlsxrb::Ooxml::XmlBuilder`**: Emits well-formed XML strings via `<<` to a writable IO, supporting streaming generation without building a DOM.
* **Part-specific parsers/writers**: `WorksheetParser`, `SharedStringsParser`, `StylesParser`, `WorkbookParser`, etc., each encapsulating the SAX event handling for one OpenXML part.

### 2. High-Level Domain Model (The "Elements" Layer)

**Namespace:** `Xlsxrb::Elements`

**Responsibility:**
This layer provides idiomatic, easy-to-use Ruby objects representing Excel concepts. It utilizes Ruby 3.2+ `Data` classes for immutability and precise structural definition. All domain models are encapsulated here to keep the top-level namespace clean.

**Core Objects (`Data` classes):**
* **`Xlsxrb::Elements::Workbook`**: Represents the entire file structure. Contains `sheets` (Array of Worksheet), shared styles metadata, and `unmapped_data`.
* **`Xlsxrb::Elements::Worksheet`**: Represents a single sheet. Contains `name`, `rows` (Array of Row), `columns` (Array of Column), and sheet-level properties.
* **`Xlsxrb::Elements::Row`**: Represents one row. Contains `index` (0-based), `cells` (Array of Cell), and row-level attributes (height, hidden, etc.).
* **`Xlsxrb::Elements::Column`**: Represents column formatting. Contains `index` (0-based), `width`, and column-level attributes.
* **`Xlsxrb::Elements::Cell`**: Represents a single cell. Contains `row_index`, `column_index` (both 0-based), `value` (Ruby native type), `formula`, `style`, and `unmapped_data`.

**Design Principles:**
* **Zero-based Indexing:** To maintain consistency with Ruby's core language (Arrays/Enumerable), all indices (rows, columns, and worksheets) are **0-based**. For Excel-style coordination, use string references like `cell("A1")`.
* **Fail-safe Design (Lazy Validation):** Exceptions are not raised during XML parsing. Each class has an `errors` property (Array of String) and a `valid?` method that returns `errors.empty?`.
* **Forward Compatibility:** All classes have an `unmapped_data` property (Hash) to ensure that any unknown XML attributes or elements are retained, preserving file integrity during round-trips.

### 3. The Facade / Entrypoint Layer

**Namespace:** `Xlsxrb` module methods

**Responsibility:**
Acts as the primary bridge, offering both In-Memory and Streaming APIs.

| Method | Type | Description |
| :--- | :--- | :--- |
| **`Xlsxrb.read(source)`** | In-Memory | Loads the entire file into a `Workbook` object. |
| **`Xlsxrb.foreach(source, **options)`** | **Streaming** | Yields each `Row` one by one. Ideal for large files. |
| **`Xlsxrb.write(target, workbook)`** | In-Memory | Saves a `Workbook` object to a file or IO. |
| **`Xlsxrb.generate(target, &block)`** | **Streaming** | Provides a DSL to stream data directly to a file/IO. |

---

## Data-Flow & Lifecycle

### `Xlsxrb.read(source)` — In-Memory Read

```
source (path / IO)
  │
  ▼
Ooxml::ZipReader          ── extracts ZIP entries ──►  raw bytes per part
  │
  ▼
Ooxml::SharedStringsParser ── SAX parse xl/sharedStrings.xml ──► string table (Array)
Ooxml::StylesParser        ── SAX parse xl/styles.xml        ──► styles hash
Ooxml::WorkbookParser      ── SAX parse xl/workbook.xml      ──► sheet list
  │
  ▼  (for each sheet)
Ooxml::WorksheetParser     ── SAX parse xl/worksheets/sheetN.xml ──►
  │                           yields (row_index, cells_array, row_attrs, unmapped)
  ▼
Elements::Cell / Row / Column / Worksheet
  │
  ▼
Elements::Workbook          ◄── assembled from all worksheets
```

### `Xlsxrb.write(target, workbook)` — In-Memory Write

```
Elements::Workbook
  │
  ▼  (for each worksheet)
Ooxml::WorksheetWriter     ── converts Row/Cell → XML fragments ──►
  │                           streams into Ooxml::ZipWriter entry
  ▼
Ooxml::WorkbookWriter      ── writes workbook.xml, styles.xml, sharedStrings.xml,
  │                           [Content_Types].xml, .rels
  ▼
Ooxml::ZipWriter           ── writes ZIP output ──► target (path / IO)
```

### `Xlsxrb.foreach(source, **options)` — Streaming Read

```
source (path / IO)
  │
  ▼
Ooxml::ZipReader           ── locates xl/sharedStrings.xml, xl/worksheets/sheetN.xml
  │
  ▼  (SAX parse SST first — kept in memory as a flat Array of strings)
  │
  ▼  (then SAX stream worksheet)
Ooxml::WorksheetParser     ── on each </row> event:
  │                           1. build Elements::Row with resolved cell values
  │                           2. yield Row to caller's block
  │                           3. discard Row (GC eligible)
  ▼
caller's block receives Elements::Row, processes, moves on
```

Key memory invariant: only **one Row** (plus the shared-string table) is alive at any time.

### `Xlsxrb.generate(target, &block)` — Streaming Write

```
caller's block
  │
  ▼  block receives a StreamWriter context object
  │  context.add_row([val1, val2, ...])
  │
  ▼
Ooxml::WorksheetWriter     ── converts array → <row><c>…</c></row> XML
  │                           writes directly to ZipWriter entry stream
  ▼
Ooxml::ZipWriter           ── compresses & writes to target
```

Key memory invariant: rows are written and flushed immediately; no row Array accumulates.

---

## Streaming Internals

### SAX-Based Worksheet Parsing (`Ooxml::WorksheetParser`)

`REXML::Parsers::SAX2Parser` fires events: `start_element`, `end_element`, `characters`.

The parser maintains a minimal state machine:

1. `start_element("row", attrs)` → record `row_index`, reset cell accumulator
2. `start_element("c", attrs)` → extract `r` (reference), `t` (type), `s` (style index)
3. `characters(text)` inside `<v>`, `<f>`, `<is>` → buffer the text
4. `end_element("c")` → build a raw cell hash, push to current row's accumulator
5. `end_element("row")` → yield the accumulated row and clear the accumulator

Any element **not** in the recognized set (`row`, `c`, `v`, `f`, `is`, `t`, etc.) is captured into `unmapped_children` on the nearest recognized ancestor.

### ZIP Streaming

`Ooxml::ZipReader` scans local file headers sequentially using `Zlib::Inflate`. It does **not** seek to the central directory — this allows reading from non-seekable IO (pipes, HTTP streams).

`Ooxml::ZipWriter` writes local file headers immediately, accumulates a central directory index in memory (entry names + offsets only), and writes the central directory + EOCD at `#close`.

---

## `unmapped_data` & Forward-Compatibility

When the Ooxml layer encounters an XML element or attribute not in its recognized set:

1. **Capture**: The element is stored as a Hash `{ tag: String, attrs: Hash, children: Array, text: String? }`.
2. **Attach**: The Hash is pushed onto the nearest recognized parent's `unmapped_children` array.
3. **Surface**: The Elements layer receives these as the `unmapped_data` field — a Hash keyed by parent-context (e.g. `{ row: [...], cell: [...], worksheet: [...] }`).
4. **Restore**: During write-back (`Ooxml::WorksheetWriter`), `unmapped_data` entries are re-serialized to XML in their original order using `XmlBuilder`, preserving any future spec extensions or vendor-specific markup.

This ensures that reading then writing an XLSX file does not silently discard unknown content.

---

## Error Handling & Validation Boundaries

### Ooxml Layer (parse-time)

* **Never raises** on unexpected XML content. Unrecognized elements → `unmapped_data`. Malformed attribute values → stored as-is (raw strings).
* **Raises** only on structural corruption that prevents further parsing (e.g., truncated ZIP, invalid UTF-8, ZIP local header CRC mismatch).

### Elements Layer (model-time)

* Each `Data` class exposes `errors` (frozen Array of String) and `valid?` (`errors.empty?`).
* Validation is performed at construction time:
  - `Cell`: value type check, column/row index range
  - `Row`: index ≥ 0, cells array consistency
  - `Worksheet`: name present, unique row indices
  - `Workbook`: at least one sheet, unique sheet names
* Invalid objects **are still created** — the caller decides how to handle `valid? == false`.

### Facade Layer

* `Xlsxrb.read` / `.foreach`: propagate Ooxml-layer structural exceptions. Content-level issues appear in `errors` on returned objects.
* `Xlsxrb.write` / `.generate`: validate the `Workbook` / row data at the boundary and raise `Xlsxrb::Error` for fatal issues (e.g., nil target path). Non-fatal issues (e.g., value truncation) are silently handled.

---

## Benefits of this Approach

* **Rubyish Interface:** Methods like `foreach` and `generate` follow Ruby's standard library conventions (e.g., `CSV.foreach`).
* **Clean Namespace:** Users only interact with the `Xlsxrb` module. Internal models are safely isolated within `Elements`.
* **Safety & LSP Support:** `Data` objects provide clear property definitions for editor autocomplete.
* **Constant Memory Streaming:** Both read and write paths support row-at-a-time processing suitable for millions of rows.
* **Future-Proofing:** The `unmapped_data` mechanism and layered design accommodate future features without rewriting the underlying XML logic.
