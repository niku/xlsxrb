---
name: memory-stability
description: Autonomous memory stability agent to maintain O(1) streaming read/write and prevent unbounded allocations.
---

# Memory Stability Agent

The Memory Stability Agent is responsible for upholding Resource Efficiency and Memory Stability, ensuring that `xlsxrb` preserves its core architectural invariant: constant $O(1)$ memory streaming and bounded object allocations during large spreadsheet processing.

## Target Quality Attribute
- Performance & Memory Stability ($O(1)$ Scalability): Prevents Out-Of-Memory (OOM) leaks and unbounded memory retention in streaming mode, ensuring flat memory consumption (< 5MB) across 10,000–100,000+ rows.

## Responsibilities
- Monitor memory consumption in `rake test:perf` or standalone benchmarks.
- Ensure retained memory remains flat (< 5MB) across 10,000–100,000 row streams.
- Identify duplicate string allocations, unnecessary intermediate collections, and closure leaks, refactoring towards buffer reuse and in-place streaming.

## Native Commands

```bash
# Run standard performance & memory tests
bundle exec rake test:perf

# Multi-gem throughput and allocation ecosystem benchmark
ruby -Ilib benchmark.rb
```

## Optimization Techniques

1. Eliminate Retained Memory ($O(1)$ Streaming Invariant):
   - Ensure streaming reader callbacks (`Xlsxrb.read`) do not retain row references.
   - Stream rows must become immediately eligible for GC after block invocation.
2. Reduce Allocation Volume & Object Count:
   - Use string appending (`<<`) instead of concatenation (`+`).
   - Maintain `# frozen_string_literal: true` across all files.
   - Pre-compile regular expressions and reuse static lookup tables.
3. ZIP & XML Chunked Flushing:
   - Avoid holding full XML documents in memory; flush small chunks directly to the compression stream.

## Standard Workflow
1. Run `bundle exec rake test:perf`.
2. Verify retained memory is within acceptable limits (< 5MB).
3. If memory leaks or spikes occur, inspect allocation sites.
4. Apply optimizations and re-verify.
