# xlsxrb Autonomous Quality Assurance (QA) Guidelines

xlsxrb is an open-source project. All modifications and contributions must adhere to the following quality gates and guidelines.

## Language Policy
As an open-source library intended for global use, all code comments, documentation, commit messages, tests, and CLI outputs must be written in English unless explicitly requested otherwise by the user.

## Writing Style & Tone
- Factual and Objective: Use restrained, engineering-focused language. Avoid promotional, hyperbolic, or vague adjectives (e.g., "robust", "comprehensive", "seamless", "enterprise-grade", "ultra-fast", "guarantee"). State actual behavior, verified metrics, and technical constraints directly.
- Minimal Markdown Formatting: Avoid ornamental bold formatting (`**...**`) across table cells, list items, or section headers. Reserve bold styling strictly for genuine emphasis or critical callout banners.

## Native Tooling Principle
To keep the repository clean and maintainable:
- Do not introduce redundant bin scripts or custom Rake tasks when existing tasks (`rake mutant:pure`, `rake typecheck`, `rake test:unit`, `rake test:perf`, `rake test:visual`, `bin/pre-push`) are already available.
- Autonomous agents directly invoke existing native tools and standard commands.

## Specialized QA Agents (Mapped to Target Quality Attributes)

1. Test Rigor Agent (`test-rigor`)
   - Quality Attribute: Test Suite Rigor / Detection Power
   - Command: `bundle exec mutant run ...` / `bundle exec rake mutant:pure`
   - Target: 100.00% kill rate (0 alive) across all pure logic and conversion subjects.

2. Type Safety Agent (`type-safety`)
   - Quality Attribute: Type Safety (Static & Dynamic)
   - Command: `bundle exec rake typecheck` / `bundle exec rake sig` / `bundle exec rake test:rbs`
   - Target: 0 type errors from Steep and full synchronization between inline annotations and `sig/generated`.

3. Specification Compliance Agent (`spec-compliance`)
   - Quality Attribute: Specification Conformance / Interoperability
   - Command: `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb` / `bundle exec rake test:e2e`
   - Target: 100% compliance with ECMA-376 Part 4 XSD schemas and Microsoft .NET Open XML SDK compatibility.

4. Memory Stability Agent (`memory-stability`)
   - Quality Attribute: Performance & Memory Stability ($O(1)$ Scalability)
   - Command: `bundle exec rake test:perf` / `ruby -Ilib benchmark.rb`
   - Target: Enforce $O(1)$ streaming memory profile (retained memory < 5MB on 10,000+ rows).

5. Throughput & Complexity Agent (`throughput-performance`)
   - Quality Attribute: Processing Speed, Allocation Efficiency & Linear Complexity ($O(N)$)
   - Command: `bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb` / `bundle exec rake benchmark:ips`
   - Target: Allocation caps (< 25 objects/row write, < 15 objects/row read), linear scaling ratio ($\le 15.0\times$ on $10\times$ scale factor).

6. Security & Resilience Agent (`security-resilience`)
   - Quality Attribute: Availability, Security & Attack Resilience
   - Command: `bundle exec ruby -Ilib -Itest test/security_resilience_test.rb` / `bundle exec rake bundle:audit`
   - Target: Zero ReDoS catastrophic backtracking, decompression bounds (`max_uncompressed_size`), safe formula escaping (`escape_formulas: true`), zero CVEs.

7. API Symmetry & Contract Agent (`api-contract`)
   - Quality Attribute: API Symmetry & Lossless Round-Trip Fidelity
   - Command: `bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb` / `bundle exec rake test:contract` / `bundle exec rake test:pbt`
   - Target: Complete method and arity symmetry between builder/writer interfaces, lossless `Read(Write(X)) = X` round-trips.

8. Concurrency & Ractor Safety Agent (`concurrency-safety`)
   - Quality Attribute: Concurrency Safety & Ractor Isolation
   - Command: `bundle exec ruby -Ilib -Itest test/concurrent_test.rb`
   - Target: 100% constant shareability across `Xlsxrb` namespace, zero class variable state, deterministic parallel Ractor pipeline execution.

9. Visual Fidelity & Wasm Compatibility Agent (`visual-fidelity`)
   - Quality Attribute: Visual Rendering Fidelity & WebAssembly Compatibility
   - Command: `bundle exec rake test:visual` / `bundle exec rake doc:check`
   - Target: Headless LibreOffice pixel match against baselines ($\le 100$px diff on tables, $\le 500$px on charts), 100% passing Wasm playground presets with COOP/COEP isolation.

## Pre-Push Verification
Before pushing changes, run:

```bash
bin/pre-push --all
```
