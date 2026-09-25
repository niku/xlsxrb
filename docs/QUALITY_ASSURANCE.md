# Quality Assurance (QA) & Testing Architecture

`xlsxrb` employs a multi-layered Quality Assurance matrix covering static code analysis, dynamic runtime validation, performance benchmarking, and visual regression testing.

Below is an overview of the verification mechanisms, their execution contexts, the quality attributes they target, and the issues they prevent.

## QA Matrix

| Inspection Mechanism / Tool | Execution Command / Mechanism | Local | CI (PR/Push) | Scheduled (Weekly) | Target Quality Attribute | Validation Method | Prevented Bugs / Issues |
| :--- | :--- | :---: | :---: | :---: | :--- | :--- | :--- |
| RuboCop | `rake rubocop` | ⭕ | ⭕ | - | Readability / Maintainability | Static Analysis (Syntax/Linting) | Overly complex methods, non-standard syntax, unused variables. |
| Steep & RBS | `rake typecheck` | ⭕ | ⭕ | - | Type Safety (Static) | Static Analysis (Type Check) | `NoMethodError`, passing incorrect arguments, method typo bugs. |
| Bundler Audit | `rake audit` | - | ⭕ | - | Security | Dependency Scanning | Inclusion of external gems with known vulnerabilities (CVEs). |
| Dependabot | `.github/dependabot.yml` | - | - | ⭕ | Currency / Maintenance | Repository Monitoring | Outdated dependencies or CI actions. |
| Unit & Contract Tests | `rake test:unit test:contract` | ⭕ | ⭕ | - | Accuracy / Functional Reqs | Dynamic Analysis (Assertions) | Method specification violations, unexpected return values, edge-case failures. |
| Runtime Type Validation (RBS::Test) | `rake test:rbs` | △ (Opt-in) | ⭕ | - | Type Safety (Dynamic) | Dynamic Analysis (Runtime Hooks) | Type errors slipping past static checks, divergence between RBS docs and implementation. |
| Property-Based Testing (PBT) | `rake test:pbt` | ⭕ | ⭕ | - | Robustness / Exhaustiveness | Automated Random Generation | Crashes caused by "unexpected inputs" (e.g., empty strings, huge numbers, special symbols like `=`). |
| Mutation Testing (Mutant) | `rake mutant:pure` | ⭕ | ⭕ | - | Test Suite Rigor / Detection Power | Fault Injection / Mutation Analysis (100% kill on 1,091 mutants) | Shallow tests, unasserted edge cases, and surviving mutant bugs in pure algorithms, predicates, and coordinates logic. |
| Security Validation (DoS Protection) | Included in `rake test:unit` | ⭕ | ⭕ | - | Availability / Safety | Dynamic Analysis (Malicious Input) | Memory/disk exhaustion from ZIP bombs, infinite parsing loops from malformed files. |
| Concurrency Validation (Thread/Ractor) | Included in `rake test:unit` | ⭕ | ⭕ | - | Thread Safety | Dynamic Analysis (Parallel Execution) | Global variable pollution, data mixing during concurrent request processing. |
| XSD Schema Validation | Included in `rake test:unit` | ⭕ | ⭕ | - | Compatibility / Compliance | Structural Validation | "We found a problem with some content in it" errors when opening in Excel. |
| E2E Interoperability Tests | `rake test:e2e` | △ (Opt-in) | ⭕ | - | Compatibility (Real-world) | 3rd-party SDK Execution | Structural defects so severe that the official .NET SDK cannot read them. |
| Memory & Performance Tests | `rake test:perf` / Action: `performance.yml` | ⭕ | ⭕ | - | Performance / Stability | Continuous Profiling (10k+ rows) | Out of Memory (OOM) leaks or unbounded memory retention in streaming mode. |
| Ecosystem Benchmarks | `ruby benchmark.rb` | ⭕ | - | - | Performance Comparison | Multi-Gem Isolated Profiling (1M cells) | Throughput and allocation comparison against competing libraries. |
| Visual Regression Testing (VRT) | `rake test:visual` | - | ⭕ | - | Visual Accuracy (UI/UX) | Headless Rendering / Pixel Diff | Visual bugs like "cell background colors dropping" or "chart layouts breaking" after code changes. |

## Autonomous QA Maintenance (Why QA Agents?)

While this QA matrix verifies multiple quality attributes, maintaining these verification layers introduces ongoing costs:
- Refactoring pure algorithms may pass all unit tests but leave surviving mutants in `rake mutant:pure` due to unasserted edge cases.
- Minor XML structural adjustments may trigger Nokogiri schema errors due to ECMA-376 `xs:sequence` definitions.
- Updating method signatures requires re-syncing inline RBS annotations (`rake sig`) and clearing Steep diagnostics.

To address this overhead and operationalize our AI-Assisted Maintenance principle, xlsxrb defines specialized agent roles and workflows ([docs/QA_AGENTS.md](QA_AGENTS.md)):

| Agent Role | Targeted Quality Attribute | Automated Remediation Strategy |
| :--- | :--- | :--- |
| Test Rigor Agent | Test Suite Rigor / Detection Power (`rake mutant:pure`) | Analyzes AST diffs of surviving mutants and adds boundary assertions until 100% kill rate is reached. |
| Type Safety Agent | Type Safety (Static & Dynamic) (`rake typecheck`, `sig/`) | Resolves type errors, updates signatures, and runs `rake sig` to align implementation with RBS definitions. |
| Specification Compliance Agent | Specification Conformance / Interoperability (`test/xsd_validation_test.rb`) | Identifies schema sequence violations and updates XML builders to follow ECMA-376 Part 4. |
| Memory Stability Agent | Resource Efficiency / Memory Stability (`rake test:perf`) | Profiles memory with MemoryProfiler to identify retention issues and preserve flat $O(1)$ streaming. |

These workflows allow contributors and AI agents alike to address test diagnostics systematically, keeping verification layers intact as the codebase evolves.

