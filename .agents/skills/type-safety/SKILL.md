---
name: type-safety
description: Autonomous type safety agent to resolve Steep and RBS errors and keep signatures in sync with implementation.
---

# Type Safety Agent

The Type Safety Agent is responsible for upholding Type Safety and Interface Consistency across `xlsxrb`, resolving static type checking errors from Steep and runtime type contract violations from `RBS::Test`.

## Target Quality Attribute
- Type Safety (Static & Dynamic): Prevents `NoMethodError`, invalid argument passings, and interface divergence by maintaining synchronization between Ruby implementation (`lib/`) and RBS signatures (`sig/` and inline annotations).

## Responsibilities
- Parse diagnostics from `bundle exec steep check` or `bundle exec rake typecheck` and `rake test:rbs`.
- Discern whether type mismatches stem from RBS signature omissions, outdated generated signatures, or implementation bugs.
- Update inline type annotations (`#: ...`) in `lib/` or handwritten signatures in `sig/`.
- Re-synchronize signatures using `bundle exec rake sig` and verify zero diagnostics.

## Native Commands

```bash
# Run Steep static type check
bundle exec steep check
# or
bundle exec rake typecheck

# Synchronize inline signatures with sig/generated
bundle exec rake sig

# Full static and runtime type contract validation
TEST_WORKERS=1 bundle exec rake test:rbs
```

## Common Type Diagnostics & Resolutions

1. Outdated Generated RBS (`sig/generated`):
   - When modifying inline annotations `#: (...) -> ...`, re-run `bundle exec rake sig`.
2. Missing Declarations (`Cannot find declaration of ...`):
   - Check `rbs_collection.yaml` or define missing types in `sig/handwritten/`.
3. Nilability (`Type X | nil is not compatible with X`):
   - Introduce guard clauses or update signatures to accept optional types (`Integer?`).
4. Union Types & Polymorphic Values:
   - Leverage existing type aliases such as `Xlsxrb::Elements::Cell::CellValue`.

## Standard Workflow
1. Run `bundle exec steep check` to view error locations.
2. Fix inline annotations in `lib/` or update implementation code.
3. Run `bundle exec rake sig && bundle exec steep check` to re-generate signatures and verify.
4. If needed, run `TEST_WORKERS=1 bundle exec rake test:rbs` to confirm runtime contracts pass.
