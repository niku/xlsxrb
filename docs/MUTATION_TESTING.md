# Mutation Testing in xlsxrb

To ensure enterprise-grade reliability and avoid "shallow test coverage" (where line coverage is high but assertions are missing or weak), `xlsxrb` adopts **Mutation Testing** via [`mbj/mutant`](https://github.com/mbj/mutant) with `test-unit` integration.

## Motivation & Strategy

Standard code coverage measures whether lines of code were executed during tests. However, it cannot guarantee that the tests will fail if the code's behavior is altered or corrupted. Mutation testing systematically modifies source code (introducing "mutants" such as flipping conditionals, substituting constants, or removing statements) and checks whether existing tests catch the mutation ("kill" the mutant).

### Targeted Scope: Pure Logic, Calculation & Algorithms

Running mutation testing across the entire codebase (which includes heavy I/O, ZIP compression, and XML streaming) is computationally expensive and prone to creating brittle tests for XML serialization. Therefore, mutation testing in `xlsxrb` is strategically focused on **pure functions, coordinates conversion, serial date calculations, and cryptographic algorithms**:

1. **Coordinates & Cell References (`Xlsxrb::Elements::Cell`, `Xlsxrb::Ooxml::Utils`)**:
   - Column letters to 0-based indices (`"A"` ↔ `0`, `"Z"` ↔ `25`, `"AA"` ↔ `26`, up to `"XFD"` ↔ `16383`).
   - Cell reference parsing (`"A1"` ↔ `[0, 0]`, range calculation).
   - Validation boundaries (`1,048,576` rows, `16,384` columns).
2. **Serial Value & DateTime Conversions (`Xlsxrb::Ooxml::Utils`)**:
   - Date ↔ 1900-based serial number conversions.
   - Time fractions and fractional days arithmetic.
   - Excel 1900 leap year bug handling.
3. **Cryptographic & Binary Parsing (`Xlsxrb::Ooxml::Crypto`, `Xlsxrb::Ooxml::Cfb`)**:
   - Agile & Standard encryption key derivation (SHA-512, AES-256, CBC padding).
   - Compound File Binary (CFB) sector and directory stream traversal.
4. **Data Integrity & Validation (`Xlsxrb::Elements::*`)**:
   - Boundary checks, constraint errors, and error aggregation.

### Excluded Scope (Non-Targets)

- **XML Serialization / Builder**: Validated via Microsoft Open XML SDK schema checks and round-trip parsing tests rather than mutant to avoid brittle assertions.
- **Visual Regression Testing (VRT)**: Heavy headless rendering with LibreOffice is run separately in CI.
- **Large Constant Lookup Tables**: Tested via property-based tests rather than copy-pasting tables.

## Running Mutation Testing

```bash
# Run mutation testing on pure logic / algorithms
bundle exec rake mutant:pure

# Or run mutant directly against a specific class or method
bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- 'Xlsxrb::Elements::Cell*'
```
