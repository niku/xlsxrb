# Mutation Testing in xlsxrb

To ensure high reliability and avoid "shallow test coverage" (where line coverage is high but assertions are missing or weak), `xlsxrb` adopts **Mutation Testing** via [`mbj/mutant`](https://github.com/mbj/mutant) with `test-unit` integration.

As of the current release, the test suite achieves **100.00% kill rate (1,091 / 1,091 mutations killed, 0 alive)** across all 38 pure functional subjects.

## Motivation & Strategy

Standard code coverage measures whether lines of code were executed during tests. However, it cannot guarantee that the tests will fail if the code's behavior is altered or corrupted. Mutation testing systematically modifies source code (introducing "mutants" such as flipping conditionals, substituting constants, or removing statements) and checks whether existing tests catch the mutation ("kill" the mutant).

### Targeted Scope: Functional Core, Pure Predicates & Byte Encoding

Running mutation testing across the entire codebase (which includes heavy file I/O, ZIP streaming, and large XML parsing) is computationally expensive and prone to creating brittle tests for streaming writers. Therefore, `xlsxrb` follows the **Functional Core / Imperative Shell** architecture, focusing mutation testing strictly on pure functions, boundary predicates, coordinate conversions, and binary encoders:

1. **Coordinates & Cell Conversions (`Xlsxrb::Elements::Cell`, `CoordinateAccess`)**:
   - Column letters to 0-based indices (`"A"` ↔ `0`, `"Z"` ↔ `25`, `"AA"` ↔ `26`, up to `"XFD"` ↔ `16383`).
   - Pure predicate checks (`Cell.valid_coordinates?`, `Cell.valid_value?`, `Cell#valid?`).
   - String, integer, and float conversions (`Cell#to_s`, `Cell#to_i`, `Cell#to_f`, `Cell#content`, `Cell#ref`).
   - Coordinate access references and cell sorting (`CoordinateAccess#cells`, `CoordinateAccess#[]`).
2. **Worksheet, Column, Row & Workbook Invariants (`Xlsxrb::Elements::*`, `Xlsxrb::StreamRow`)**:
   - Worksheet name validation (`Worksheet.valid_name?` with 1..31 characters and forbidden character rules `\ / ? * [ ]`).
   - Column and row bounds (`Column.valid_index?`, `Column.validate`, `Column#valid?`, `Row.valid_index?`, `Row#valid?`).
   - Workbook structure validation (`Workbook#valid?`).
   - Streaming row data integrity (`StreamRow#valid?`, `StreamRow#unmapped_data`, `StreamRow#errors`, `StreamRow#values`).
3. **Serial Value & DateTime Conversions (`Xlsxrb::Ooxml::Utils`)**:
   - Julian Day ↔ 1900-based serial number conversions (`date_to_serial`, `serial_to_date`).
   - Time fractions and fractional days arithmetic (`datetime_to_serial`, `serial_to_datetime`).
   - Excel 1900 leap year bug handling.
4. **Binary Parsing & Bitfield Packings (`Xlsxrb::Ooxml::ZipGenerator`, `Xlsxrb::Ooxml::Cfb`)**:
   - Little-endian 16/32-bit byte serialization (`ZipGenerator.le16`, `ZipGenerator.le32`).
   - MS-DOS packed datetime calculation with arithmetic bitfield combinations (`ZipGenerator.dos_datetime`).
   - Compound File Binary (CFB) magic header recognition (`Cfb::Reader.cfb?`) and directory entry predicates (`DirEntry#stream?`, `#root?`, `#storage?`).
5. **XML Character Escaping (`Xlsxrb::Ooxml::XmlBuilder`)**:
   - Pure escaping of XML special characters (`XmlBuilder.escape`) with object identity preservation for clean strings.
6. **DSL Normalization (`Xlsxrb::DslHelpers`)**:
   - Column indices, range, and alphabet normalization (`DslHelpers.normalize_column_indices`).

### Excluded Scope (Non-Targets)

- **Streaming XML Writers & Parsers**: Validated via Microsoft Open XML SDK schema checks and round-trip parsing tests rather than mutant to avoid brittle assertions.
- **Visual Regression Testing (VRT)**: Headless rendering with LibreOffice Calc is validated via image diffs in CI.
- **Cryptographic Key Stretching**: Password hashing with 100,000 spin iterations is validated by specific contract tests to avoid slow feedback cycles.

## Running Mutation Testing

```bash
# Run full mutation testing suite (38 subjects, ~35 seconds)
bundle exec rake mutant:pure

# Or run mutant against a specific class or method using .mutant.yml configuration
bundle exec mutant run -- 'Xlsxrb::Elements::Cell.valid_value?'
```
