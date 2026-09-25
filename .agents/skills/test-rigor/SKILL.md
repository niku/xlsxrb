---
name: test-rigor
description: Autonomous test rigor agent to eliminate surviving mutants and maintain a 100% kill rate on pure logic in xlsxrb.
---

# Test Rigor Agent

The Test Rigor Agent is responsible for upholding Test Suite Rigor and Fault Detection Power across all pure logic and algorithms in `xlsxrb` (coordinate conversions, serial dates, binary packaging, XML escaping, row/column boundary invariants) using mutation testing (`mbj/mutant`).

## Target Quality Attribute
- Test Suite Rigor / Detection Power: Ensures tests do not merely execute lines (C0 coverage), but assert expected behaviors so that semantic mutations are detected and killed.

## Responsibilities
- Identify surviving mutants (`alive > 0`) using `bundle exec mutant run` or `bundle exec rake mutant:pure`.
- Analyze mutation diffs (AST alterations, inverted conditionals, boundary value adjustments, nil replacements).
- Add targeted, meaningful assertions in `test/` verifying the precise boundary behavior or invariant.
- Iterate autonomously until a 100.00% kill rate (0 alive) is restored across all 38 pure functional subjects.

## Native Commands

```bash
# Test a specific method or subject directly
bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell.valid_value?"
bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell#to_i"

# Test all 38 pure subjects (~28s)
bundle exec rake mutant:pure
```

## Mutation Analysis Patterns

1. Inverted Conditions (`if a` -> `if !a`):
   - Add explicit assertions for the false branch and error handling behaviors.
2. Boundary Operator Mutations (`> 0` -> `>= 0`, `< 31` -> `<= 31`):
   - Test exact edge boundaries (e.g., length 0, 1, 31, 32; indices -1, 0, 16383, 16384).
3. Nil Injections (`return x` -> `return nil`):
   - Assert that returned values are non-nil and match the expected type and content.
4. Operator Mutations (`+` -> `-`, `*` -> `/`):
   - Use non-trivial test inputs where operations do not accidentally produce identical values.
5. Removed Exceptions (dropping `raise`):
   - Verify that invalid inputs raise `ArgumentError` or expected exception classes.

## Standard Workflow
1. Execute `bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "<subject>"`.
2. Inspect the surviving mutation diff (`- original`, `+ mutant`).
3. Add targeted test assertions to the corresponding test file in `test/`.
4. Run `bundle exec rake test:unit` to verify tests pass.
5. Re-run the mutant command to confirm 100% kill rate.
