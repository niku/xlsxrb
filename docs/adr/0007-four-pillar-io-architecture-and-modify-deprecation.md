# ADR 0007: Four-Pillar I/O Architecture and In-Place XML Modification Deprecation

* Status: Accepted
* Date: 2026-11-06
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb`](../../lib/xlsxrb.rb), [`Xlsxrb::Elements::Workbook`](../../lib/xlsxrb/elements/workbook.rb), [`Xlsxrb::Elements::Worksheet`](../../lib/xlsxrb/elements/worksheet.rb), [`Xlsxrb::StreamSheet`](../../lib/xlsxrb/stream_sheet.rb)

---

## Context

xlsxrb originally offered three primary entry points: `Xlsxrb.read`, `Xlsxrb.write`, and `Xlsxrb.modify`. Over time, two architectural issues emerged:

1. Ambiguous Read Lifecycles and Resource Leaks:
   `Xlsxrb.read` behaved conditionally:
   * When given a block, `Xlsxrb.read(path) { |sheet| ... }` yielded a streaming `StreamSheet`, parsing rows on demand in $O(1)$ memory and closing the underlying ZIP container upon block termination.
   * Without a block, `wb = Xlsxrb.read(path)` returned a lazily evaluated `Workbook` holding open file descriptors and streaming sheet proxies. Callers expecting an in-memory workbook were forced to call `wb.sheet(0).load` or `wb.load`. This duality polluted static type signatures (`Elements::Worksheet | StreamSheet`) and created silent file-descriptor leaks in long-running web servers (Puma/Sidekiq) when callers forgot to call `wb.close`.

2. Fragility and Limited Utility of In-Place Modification (`modify`):
   `Xlsxrb.modify(path, output_path)` attempted in-memory spreadsheet modification by parsing cells, modifying values, and regex/byte-slice patching the raw XML inside the ZIP. In practice:
   * It only worked for small in-memory sheets, failing completely for large spreadsheets requiring $O(1)$ streaming.
   * Patching `<sheetData>` XML cannot safely preserve formulas, shared strings, calculation chains, cell styles, conditional formatting, and drawing relationships without re-serializing the entire OpenXML package.
   * In modern workflows, template-based document generation is more reliably performed by code generation (reading template structures with `Xlsxrb.load` and generating declarative `Xlsxrb.write` DSL code), while runtime data transformation is handled by streaming pipelines (`Xlsxrb.read` -> `Xlsxrb.write`).

---

## Guiding Principles

1. 2x2 Symmetric Facade Matrix:
   I/O entry points must explicitly distinguish between $O(1)$ streaming pipelines and in-memory random-access documents across both read and write axes.
2. Safe Lifecycle by Default:
   Opening a file without a block must not leave unmanaged open file descriptors or lazily coupled ZIP readers.
3. Separation of ETL Transformations from Document Construction:
   Transformations must stream rows from reader to writer rather than mutating intermediate XML buffers in place.

---

## Decision

### 1. Four-Pillar Facade Architecture
Establish four canonical entry points forming a 2x2 matrix:

| Axis | Streaming ($O(1)$ Memory) | In-Memory (Full Random Access) |
| :--- | :--- | :--- |
| Read / Input | `Xlsxrb.read(source) { \|sheet\| ... }` | `Xlsxrb.load(source, **opts) -> Elements::Workbook` |
| Write / Output | `Xlsxrb.write(dest, **opts) { \|wb\| ... }` | `Xlsxrb.build(**opts) { \|wb\| ... } -> Elements::Workbook` |

* `Xlsxrb.load(source, **opts)`: Immediately parses the entire workbook and its sheets into in-memory `Elements::Workbook` and `Elements::Worksheet` objects, and completely closes the underlying ZIP reader before returning.
* Blockless `Xlsxrb.read(source)` and `Workbook#load`: Removed in 0.2.0. Calling `Xlsxrb.read` without a block returns an `Enumerator` for streaming sheets; for in-memory random access, callers must use `Xlsxrb.load(source)`.
* `StreamSheet#load`: Retained as a public, canonical method for single-sheet materialization (e.g. selectively loading one heavy sheet into memory during streaming while leaving remaining sheets streamed).

### 2. Deprecation of Xlsxrb.modify and In-Place Mutations
* `Xlsxrb.modify`, `Workbook#update_sheet`, and `Worksheet#update_cell` were deprecated in 0.1.x and completely removed in 0.2.0.
* Direct runtime transformations to the streaming pipeline:
  ```ruby
  Xlsxrb.write(output_path) do |wb|
    wb.sheet("Processed") do |out_sheet|
      Xlsxrb.read(input_path) do |in_sheet|
        in_sheet.each_row do |row|
          out_sheet.row(row.values(type_cast: true).map { |v| v.is_a?(Numeric) ? v * 1.1 : v })
        end
      end
    end
  end
  ```
* Guide template migration: Use `Xlsxrb.load` with LLM coding agents to convert legacy Excel templates into reproducible, declarative `Xlsxrb.write` DSL code.

### 3. Serialization Facade Consolidation
* `Xlsxrb.write_workbook` and `Xlsxrb.workbook_to_data` are deprecated in favor of `Xlsxrb.write(target, workbook)` and `Xlsxrb.write(workbook)`.
* In-memory serialization helpers `Elements::Workbook#save` and `Elements::Workbook#to_data` delegate directly to `Xlsxrb.write` without emitting deprecation warnings.

### 4. On-Demand ZIP Reader Reopening
* `Ooxml::ZipReader` tracks the origin file path when initialized from a file path source. If entry extraction (`read_entry` or `each_entry_chunk`) occurs after an initial stream close, the reader reopens the file on demand to fulfill the read without raising `IOError: closed stream`.

### 5. Private Low-Level OOXML Tooling
* Low-level OpenXML components `Ooxml::Reader`, `Ooxml::Writer`, and `Ooxml::ZipGenerator` are marked `# @api private` to direct users to canonical facades (`Xlsxrb.load`, `Xlsxrb.read`, `Xlsxrb.write`, `Xlsxrb.build`).

---

## Consequences

### Positive
* Resource Leak Elimination: `Xlsxrb.load` guarantees all file handles and ZIP archives are closed before returning to user code.
* Type Safety: Removes union types (`Worksheet | StreamSheet`). `Xlsxrb.load` always returns `Elements::Workbook` with `Elements::Worksheet` instances; `Xlsxrb.read` always yields `StreamSheet` instances.
* Architectural Simplicity: Removes fragile regex-based XML patching logic from the codebase.

### Negative / Trade-offs
* Migration: Existing scripts using `Xlsxrb.modify` must transition to streaming pipelines (`Xlsxrb.read` + `Xlsxrb.write`) or `Xlsxrb.load` + `Xlsxrb.write`. Clear deprecation warnings and documentation facilitate this upgrade.
