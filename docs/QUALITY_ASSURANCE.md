# Quality Assurance (QA) & Testing Architecture

`xlsxrb` employs a multi-layered Quality Assurance matrix covering static code analysis, dynamic runtime validation, performance benchmarking, security resilience, concurrency verification, and visual regression testing.

Below is an overview of the verification mechanisms, their execution contexts, the quality attributes they target, and the issues they prevent.

## QA Matrix

| Inspection Mechanism / Tool | Execution Command / Mechanism | Local | CI (PR/Push) | Scheduled (Weekly) | Target Quality Attribute | Validation Method | Prevented Bugs / Issues |
| :--- | :--- | :---: | :---: | :---: | :--- | :--- | :--- |
| RuboCop | `rake rubocop` | ⭕ | ⭕ | - | Readability / Maintainability | Static Analysis (Syntax/Linting) | Overly complex methods, non-standard syntax, unused variables. |
| Steep & RBS | `rake typecheck` | ⭕ | ⭕ | - | Type Safety (Static) | Static Analysis (Type Check) | `NoMethodError`, passing incorrect arguments, method typo bugs. |
| Bundler Audit | `rake bundle:audit` | ⭕ | ⭕ | - | Security | Dependency Scanning | Inclusion of external gems with known vulnerabilities (CVEs). |
| Dependabot | `.github/dependabot.yml` | - | - | ⭕ | Currency / Maintenance | Repository Monitoring | Outdated dependencies or CI actions. |
| Unit & Contract Tests | `rake test:unit test:contract` | ⭕ | ⭕ | - | Accuracy / Functional Reqs | Dynamic Analysis (Assertions) | Method specification violations, unexpected return values, edge-case failures. |
| Runtime Type Validation (RBS::Test) | `rake test:rbs` | △ (Opt-in) | ⭕ | - | Type Safety (Dynamic) | Dynamic Analysis (Runtime Hooks) | Type errors slipping past static checks, divergence between RBS docs and implementation. |
| Property-Based Testing (PBT) | `rake test:pbt` | ⭕ | ⭕ | - | Robustness / Exhaustiveness | Automated Random Generation | Crashes caused by "unexpected inputs" (e.g., empty strings, huge numbers, special symbols like `=`). |
| API Symmetry Verification | `ruby -Ilib -Itest test/contract/api_symmetry_test.rb` | ⭕ | ⭕ | - | API Symmetry / Consistency | Reflection / Signature Comparison | Feature drift where methods are added to in-memory builder but omitted from streaming writer. |
| Mutation Testing (Mutant) | `rake mutant:pure` | ⭕ | ⭕ | - | Test Suite Rigor / Detection Power | Fault Injection / Mutation Analysis (100% kill on 1,528 mutants across 43 subjects) | Shallow tests, unasserted edge cases, and surviving mutant bugs in pure algorithms, predicates, and coordinates logic. |
| Security Validation (DoS & ReDoS) | `ruby -Ilib -Itest test/security_resilience_test.rb test/dos_protection_test.rb test/zip_bomb_test.rb` | ⭕ | ⭕ | - | Availability / Attack Resilience | Adversarial Fuzzing / Timing Checks | ReDoS timeouts, memory exhaustion from ZIP bombs, entity expansion DoS, formula injection. |
| Concurrency Validation (Thread/Ractor) | `ruby -Ilib -Itest test/concurrent_test.rb` | ⭕ | ⭕ | - | Concurrency Safety & Ractor Isolation | Reflection Scan & Parallel Execution | Unshareable mutable constants, global state corruption, `Ractor::IsolationError`. |
| XSD Schema Validation | `ruby -Ilib -Itest test/xsd_validation_test.rb` | ⭕ | ⭕ | - | Compatibility / Compliance | Structural Validation | "We found a problem with some content in it" errors when opening in Excel. |
| E2E Interoperability Tests | `rake test:e2e` | △ (Opt-in) | ⭕ | - | Compatibility (Real-world) | 3rd-party SDK Execution | Structural defects so severe that the official .NET SDK cannot read them. |
| Memory Stability Tests | `rake test:perf` / Action: `performance.yml` | ⭕ | ⭕ | - | Resource Efficiency ($O(1)$ Memory) | Continuous Profiling (10k+ rows) | Out of Memory (OOM) leaks or unbounded memory retention in streaming mode. |
| Throughput & Complexity Tests | `ruby -Ilib -Itest test/performance/throughput_test.rb` | ⭕ | ⭕ | - | Processing Speed & Linear Scaling ($O(N)$) | Allocation Counting & Scaling Ratios | Quadratic runtime degradation ($O(N^2)$), excessive object allocations per row. |
| Hot-Path Microbenchmarks | `rake benchmark:ips` | ⭕ | - | - | Performance Optimization | Iterations Per Second (benchmark-ips) | Latency regressions in XML escaping, cell coordinate math, and row parsing loops. |
| Ecosystem Benchmarks | `ruby benchmark.rb` | ⭕ | - | - | Ecosystem Comparative Performance | Multi-Gem Isolated Profiling (1M cells) | Throughput and allocation comparison against competing libraries. |
| Visual Regression Testing (VRT) | `rake test:visual` | - | ⭕ | - | Visual Accuracy (UI/UX) | Headless Rendering / Pixel Diff | Visual bugs like "cell background colors dropping" or "chart layouts breaking" after code changes. |
| Wasm Playground Validation | `rake doc:check` | ⭕ | ⭕ | - | WebAssembly Runtime Compatibility | Static & Dynamic Asset Verification | Browser execution crashes, missing COOP/COEP isolation headers, broken preset examples. |

## Autonomous QA Maintenance (Why QA Agents?)

While this QA matrix verifies multiple quality attributes, maintaining these verification layers introduces ongoing costs:
- Refactoring pure algorithms may pass all unit tests but leave surviving mutants in `rake mutant:pure` due to unasserted edge cases.
- Minor XML structural adjustments may trigger Nokogiri schema errors due to ECMA-376 `xs:sequence` definitions.
- Updating method signatures requires re-syncing inline RBS annotations (`rake sig`) and clearing Steep diagnostics.
- Adding features to `WorksheetBuilder` without updating `StreamWriter` causes architectural divergence.
- Introducing un-frozen constants breaks multi-Ractor parallel workloads.

To address this overhead and operationalize our AI-Assisted Maintenance principle, xlsxrb defines specialized agent roles and workflows ([docs/QA_AGENTS.md](QA_AGENTS.md)):

| Agent Role | Targeted Quality Attribute | Automated Remediation Strategy |
| :--- | :--- | :--- |
| Test Rigor Agent | Test Suite Rigor / Detection Power (`rake mutant:pure`) | Analyzes AST diffs of surviving mutants and adds boundary assertions until 100% kill rate is reached. |
| Type Safety Agent | Type Safety (Static & Dynamic) (`rake typecheck`, `sig/`) | Resolves type errors, updates signatures, and runs `rake sig` to align implementation with RBS definitions. |
| Specification Compliance Agent | Specification Conformance / Interoperability (`test/xsd_validation_test.rb`) | Identifies schema sequence violations and updates XML builders to follow ECMA-376 Part 4. |
| Memory Stability Agent | Resource Efficiency / Memory Stability (`rake test:perf`) | Profiles memory with MemoryProfiler to identify retention issues and preserve flat $O(1)$ streaming. |
| Throughput & Complexity Agent | Processing Speed & Linear Scaling (`test/performance/throughput_test.rb`, `rake benchmark:ips`) | Verifies per-row allocation caps and $O(N)$ scaling ratios; benchmarks hot paths with `benchmark-ips`. |
| Security & Resilience Agent | Attack Resilience & DoS Defense (`test/security_resilience_test.rb`, `rake bundle:audit`) | Audits regexes for ReDoS with 10k+ character payloads, guards ZIP bomb limits, and checks dependencies. |
| API Symmetry & Contract Agent | API Symmetry & Round-Trip Fidelity (`test/contract/api_symmetry_test.rb`, `rake test:pbt`) | Asserts reflective method parity between builders and verifies lossless roundtrips with property-based testing. |
| Concurrency & Ractor Safety Agent | Thread Safety & Ractor Isolation (`test/concurrent_test.rb`) | Scans all namespace constants for `Ractor.shareable?` and executes parallel write/read/modify pipelines. |
| Visual Fidelity & Wasm Agent | Visual Fidelity & Wasm Compatibility (`rake test:visual`, `rake doc:check`) | Compares LibreOffice Calc rendering against pixel baselines and validates browser Wasm playground presets. |

These workflows allow contributors and AI agents alike to address test diagnostics systematically, keeping verification layers intact as the codebase evolves.
