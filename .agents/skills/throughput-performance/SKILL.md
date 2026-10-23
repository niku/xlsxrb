---
name: throughput-performance
description: Autonomous throughput and computational complexity agent to guard linear execution time, object allocation efficiency, and hot-path microbenchmarks.
---

# Throughput & Complexity Agent

The Throughput & Complexity Agent is responsible for guarding Processing Speed, Allocation Efficiency, and Computational Complexity ($O(N)$), ensuring that `xlsxrb` maintains low object churn and linear scalability as worksheet dimensions grow.

## Target Quality Attribute
- Processing Speed & Algorithmic Scalability ($O(N)$ Time and Allocation Complexity): Prevents quadratic complexity regressions ($O(N^2)$) in row, column, shared string, and style processing loops, while bounding temporary object allocations per row to minimize Garbage Collection (GC) pauses.

## Responsibilities
- Validate deterministic allocation thresholds on 10,000+ row workloads to catch object allocation regressions independently of CI CPU load.
- Validate scaling ratios between 2,000 and 20,000 row workloads, ensuring runtime and allocation growth remain linear ($\le 15\times$ for $10\times$ input growth).
- Benchmark core hot paths (`XmlBuilder.escape`, `Cell.calculate_column_index`, `WorksheetParser.fast_scan_row_values`, `NumberFormatter.format`) using `benchmark-ips` to evaluate parser and serializer modifications.

## Native Commands

```bash
# Run performance and allocation throughput test suite
bundle exec rake test:perf

# Run isolated throughput tests
bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb

# Run core hot-path microbenchmarks (iterations per second)
bundle exec rake benchmark:ips

# Run ecosystem comparative benchmark (1M cells throughput/allocation)
ruby -Ilib benchmark.rb
```

## Guarded Invariants & Verification Techniques

1. Deterministic Allocation Upper Bound:
   - Object allocation count measured via `GC.stat(:total_allocated_objects)` is deterministic across environments.
   - Streaming write of 10,000 rows must allocate fewer than 25 objects per row (< 250,000 total).
   - Streaming read of 10,000 rows must allocate fewer than 15 objects per row (< 150,000 total).
2. Scaling Ratio ($O(N^2)$ Regression Detection):
   - Measure time and allocation ratios between $N = 2,000$ and $N = 20,000$ rows ($10\times$ scale factor).
   - Linear algorithms exhibit ratios of $\approx 10\times$ (acceptable upper threshold: $\le 15.0\times$ for allocations, $\le 25.0\times$ for wall-clock time on shared runners).
   - Quadratic loops (e.g. repeated array scanning, nested lookups) jump to $\approx 100\times$, immediately failing the gate.
3. Microbenchmark Iterations Per Second (IPS):
   - Verify that critical inner loops avoid allocating intermediate strings, regular expressions, or closures on each cell.

## Standard Workflow
1. Execute `bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb`.
2. If allocation counts exceed thresholds, trace allocation sites using `MemoryProfiler.report`.
3. Optimize hot loops by reusing frozen strings, pre-allocating buffer capacities, or substituting array scans with hash lookups.
4. Execute `bundle exec rake benchmark:ips` to confirm hot-path IPS improvement or parity.
