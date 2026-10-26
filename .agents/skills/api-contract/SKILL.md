---
name: api-contract
description: Autonomous API symmetry and round-trip contract agent to prevent feature divergence between in-memory and streaming builders and enforce lossless serialization.
---

# API Symmetry & Contract Agent

The API Symmetry & Contract Agent is responsible for guarding Interface Consistency and Lossless Round-Trip Fidelity, ensuring that public worksheet authoring capabilities and query methods remain strictly symmetrical between In-Memory (`Xlsxrb.build`, `Elements::Worksheet`) and Streaming (`Xlsxrb.write`, `StreamWriter`, `StreamSheet`) architectures.

## Target Quality Attribute
- API Symmetry & Round-Trip Completeness: Eliminates interface drift between streaming and in-memory builders, and ensures that all cell data types, styles, formulas, and worksheet attributes survive round-trip serialization without data loss (`Read(Write(X)) = X`).

## Responsibilities
- Validate reflective method signature symmetry: assert that every authoring method on `WorkbookBuilder` and `WorksheetBuilder` is implemented with compatible positional and keyword parameter signatures on `StreamWriter` and `StreamWriter::WorksheetProxy`.
- Validate query method symmetry between in-memory `Elements::Worksheet` and streaming `StreamSheet` (`[]`, `cell`, `each`, `name`, `state`, `hidden?`, `visible?`, `very_hidden?`).
- Run the contract test suite (`rake test:contract`) to verify semantic and structural equivalence of generated OOXML (such as `<definedNames>` scoping and emission order) across streaming and in-memory paths.
- Run property-based tests (`rake test:pbt`) to fuzz random union-type inputs and special Unicode/formula boundary sequences.

## Native Commands

```bash
# Run reflective API symmetry tests
bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb

# Run contract tests verifying semantic parity across APIs
bundle exec rake test:contract

# Run property-based roundtrip testing suite
bundle exec rake test:pbt
```

## Architectural Invariants

1. Builder Symmetry Invariant:
   - Any new workbook, worksheet, or cell authoring capability added to `WorkbookBuilder` or `WorksheetBuilder` must be mirrored in `StreamWriter` and delegated via `WorksheetProxy`.
   - Method arities, optional parameters, and keyword parameter names (`:key`, `:keyreq`) must match across builder interfaces.
2. Reader Query Interface Invariant:
   - In-memory `Elements::Worksheet` and streaming `StreamSheet` must support identical basic query predicates and coordinate access methods (`[]`, `cell`, `state`, `hidden?`, `visible?`, `very_hidden?`).
3. Lossless Round-Trip Invariant:
   - Writing data through either streaming or in-memory API and reading back via either streaming or in-memory parser must yield byte-for-byte or value-for-value identical representations.
4. Normalization Idempotency Invariant:
   - Shared DSL normalizers in `DslHelpers` (e.g., `normalize_defined_name_value`, `format_absolute_reference`, `normalize_data_validation`) must be idempotent ($f(f(x)) = f(x)$).
   - When a worksheet-level helper normalizes coordinates and delegates to a workbook-level method, the workbook-level normalizer must leave already-qualified references (containing `!`), formulas (starting with `=`), and numeric constants untouched to prevent double-quoting or corruption.
5. OOXML Scoping & Chronological Order Invariant:
   - Worksheet-scoped definitions (e.g., `sheet.defined_name` setting `local_sheet_id` to the 0-based sheet index) and workbook-scoped definitions (`wb.defined_name` with `local_sheet_id: nil`) must be emitted in the exact same chronological call order in both `WorkbookBuilder` and `StreamWriter`.

## Standard Workflow
1. When introducing a new workbook or worksheet authoring option or DSL method, implement normalization once in `DslHelpers` and wire it identically into `WorkbookBuilder` / `WorksheetBuilder` and `StreamWriter` / `WorksheetProxy`.
2. Run `bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb` to confirm zero method, keyword, or XML parity discrepancies.
3. Run `bundle exec rake test:contract` and `bundle exec rake test:pbt` to verify round-trip fidelity.
