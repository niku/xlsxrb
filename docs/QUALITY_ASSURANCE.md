# Quality Assurance (QA) & Testing Architecture

`xlsxrb` is designed to be an enterprise-grade, highly reliable, and highly performant library for reading and writing Excel spreadsheets. To achieve and maintain this standard, we have implemented a comprehensive Quality Assurance matrix that covers everything from static code analysis and dynamic runtime validation to performance benchmarking and visual regression testing.

Below is an overview of the inspection mechanisms, when they run, the quality attributes they guarantee, and the specific bugs they prevent.

## QA Matrix

| Inspection Mechanism / Tool | Execution Command / Mechanism | Local | CI (PR/Push) | Scheduled (Weekly) | Target Quality Attribute | Validation Method | Prevented Bugs / Issues |
| :--- | :--- | :---: | :---: | :---: | :--- | :--- | :--- |
| **RuboCop** | `rake rubocop` | ⭕ | ⭕ | - | Readability / Maintainability | Static Analysis (Syntax/Linting) | Overly complex methods, non-standard syntax, unused variables. |
| **Steep & RBS** | `rake typecheck` | ⭕ | ⭕ | - | Type Safety (Static) | Static Analysis (Type Check) | `NoMethodError`, passing incorrect arguments, method typo bugs. |
| **Bundler Audit** | `rake audit` | - | ⭕ | - | Security | Dependency Scanning | Inclusion of external gems with known vulnerabilities (CVEs). |
| **Dependabot** | `.github/dependabot.yml` | - | - | ⭕ | Currency / Maintenance | Repository Monitoring | Outdated dependencies or CI actions. |
| **Unit & Contract Tests** | `rake test:unit test:contract` | ⭕ | ⭕ | - | Accuracy / Functional Reqs | Dynamic Analysis (Assertions) | Method specification violations, unexpected return values, edge-case failures. |
| **Runtime Type Validation (RBS::Test)** | `rake test:rbs` | △ (Opt-in) | ⭕ | - | Type Safety (Dynamic) | Dynamic Analysis (Runtime Hooks) | Type errors slipping past static checks, divergence between RBS docs and implementation. |
| **Property-Based Testing (PBT)** | Included in `rake test:unit` | ⭕ | ⭕ | - | Robustness / Exhaustiveness | Automated Random Generation | Crashes caused by "unexpected inputs" (e.g., empty strings, huge numbers, special symbols like `=`). |

| **Security Validation (DoS Protection)** | Included in `rake test:unit` | ⭕ | ⭕ | - | Availability / Safety | Dynamic Analysis (Malicious Input) | Memory/disk exhaustion from ZIP bombs, infinite parsing loops from malformed files. |
| **Concurrency Validation (Thread/Ractor)** | Included in `rake test:unit` | ⭕ | ⭕ | - | Thread Safety | Dynamic Analysis (Parallel Execution) | Global variable pollution, data mixing during concurrent request processing. |
| **XSD Schema Validation** | Included in `rake test:unit` | ⭕ | ⭕ | - | Compatibility / Compliance | Structural Validation | "We found a problem with some content in it" errors when opening in Excel. |
| **E2E Interoperability Tests** | `rake test:e2e` | △ (Opt-in) | ⭕ | - | Compatibility (Real-world) | 3rd-party SDK Execution | Structural defects so severe that the official .NET SDK cannot read them. |
| **Load / Stress Testing** | Action: `benchmark.yml` | - | ⭕ | - | Performance / Stability | Massive Data Generation (e.g., 10k+ rows) | Out of Memory (OOM) crashes or extreme delays when processing large datasets. |
| **Memory & Speed Benchmark** | Action: `benchmark.yml` | - | ⭕ | - | Performance | Continuous Profiling | Memory leaks, severe performance degradation due to inefficient loop additions. |
| **Visual Regression Testing (VRT)** | `rake test:visual` | - | ⭕ | - | Visual Accuracy (UI/UX) | Headless Rendering / Pixel Diff | Visual bugs like "cell background colors dropping" or "chart layouts breaking" after code changes. |

