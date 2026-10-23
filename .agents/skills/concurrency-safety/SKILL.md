---
name: concurrency-safety
description: Autonomous concurrency and Ractor safety agent to guard thread safety, constant shareability, and isolation in parallel Ruby 4.0+ execution.
---

# Concurrency & Ractor Safety Agent

The Concurrency & Ractor Safety Agent is responsible for guarding Thread Safety, Ractor Isolation, and Parallel Scalability, ensuring that `xlsxrb` operates without state corruption, data mixing, or `Ractor::IsolationError` across multi-threaded web servers and parallel Ractor pipelines.

## Target Quality Attribute
- Concurrency Safety & Ractor Isolation: Guarantees zero global mutable state, strict shareability of all library constants, and safe parallel execution across threads and Ractors.

## Responsibilities
- Recursively verify that all constants defined within the `Xlsxrb` namespace satisfy `Ractor.shareable?(val) == true`, eliminating un-frozen lookup tables, string literals, arrays, and date epochs.
- Enforce clean separation of execution state, preventing class variables (`@@`) or class instance variables from being accessed across Ractors.
- Verify full-pipeline Ractor parallelism (`test_full_pipeline_ractor_parallelism`): execute concurrent XLSX generation, streaming parsing, in-memory loading, and template modification within isolated Ractors without synchronization failures.

## Native Commands

```bash
# Run concurrency and Ractor safety test suite
bundle exec ruby -Ilib -Itest test/concurrent_test.rb
```

## Concurrency Invariants

1. Ractor Shareability Invariant:
   - All collection constants (arrays, hashes, sets) and custom object constants (Date epochs, byte strings) must be deeply frozen or made shareable with `Ractor.make_shareable` or `.freeze`.
   - Any mutable constant triggers a test failure in `test_all_xlsxrb_constants_are_ractor_shareable`.
2. No Class Variable (`@@`) State:
   - In Ruby 4.0+, reading or writing class variables from non-main Ractors triggers a fatal `Ractor::IsolationError`.
   - Maintain Ractor-compatible accessors for third-party or stdlib dependencies (such as REXML security settings).
3. Deterministic Parallel Execution:
   - Concurrent calls to `Xlsxrb.write`, `Xlsxrb.read`, or `Xlsxrb.modify` on distinct IO streams must produce strictly deterministic results without cross-worker interference.

## Standard Workflow
1. When declaring new lookup tables, default option hashes, format masks, or regex constants in `lib/`, ensure they are frozen and deeply shareable.
2. Run `bundle exec ruby -Ilib -Itest test/concurrent_test.rb`.
3. If an unshareable constant or isolation error is identified, freeze the nested objects or apply `Ractor.make_shareable`.
4. Re-verify the full pipeline under parallel Ractor execution.
