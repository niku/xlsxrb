# Autonomous Quality Assurance (QA) Agents in xlsxrb

This document defines the architecture, responsibilities, standard commands, and autonomous workflows of the Autonomous QA Agents Squad for xlsxrb.

---

## 1. Overview & Architectural Philosophy

xlsxrb is a pure Ruby library for reading and writing Excel spreadsheets with streaming support. Rather than introducing wrapper scripts or ad-hoc tasks, the autonomous QA agents directly leverage the repository's native verification infrastructure to guard specific quality attributes:

```
┌────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
│                                                xlsxrb Autonomous QA Squad                                              │
├────────────────────┬────────────────────┬────────────────────┬────────────────────┬────────────────────────────────────┤
│     Test Rigor     │    Type Safety     │  Spec Compliance   │  Memory Stability  │       Throughput & Complexity      │
│ (Mutant 100% Kill) │   (Steep & RBS)    │  (ECMA-376 SDK)    │   (O(1) Memory)    │        (O(N) & Low GC Churn)       │
├────────────────────┼────────────────────┼────────────────────┼────────────────────┼────────────────────────────────────┤
│ bundle exec        │ bundle exec rake   │ bundle exec ruby   │ bundle exec rake   │ bundle exec rake test:perf         │
│ mutant:pure        │ typecheck / sig    │ test/xsd_...       │ test:perf          │ bundle exec rake benchmark:ips     │
├────────────────────┼────────────────────┼────────────────────┼────────────────────┴────────────────────────────────────┤
│Security&Resilience │    API Contract    │ Concurrency Safety │  Visual Fidelity & Wasm Compatibility                  │
│(ReDoS/DoS/Crypto)  │(Symmetry & Round)  │(Ractor Isolation)  │  (LibreOffice Pixel VRT & ruby.wasm Playground)        │
├────────────────────┼────────────────────┼────────────────────┼─────────────────────────────────────────────────────────┤
│ bundle exec ruby   │ bundle exec rake   │ bundle exec ruby   │ bundle exec rake test:visual                            │
│ test/security_...  │ test:contract/pbt  │ test/concurrent... │ bundle exec rake doc:check                             │
└────────────────────┴────────────────────┴────────────────────┴─────────────────────────────────────────────────────────┘
```

---

## 2. Specialized Agents & Quality Attributes

### 1. Test Rigor Agent (`test-rigor`)
- Target Quality Attribute: Test Suite Rigor / Detection Power
- Mission: Eliminate surviving mutants in pure logic, algorithms, and conversions to maintain a 100.00% kill rate across all 43 pure functional subjects.
- Scope: Coordinate conversions, Julian day/serial dates, binary packaging, XML escaping, row/column boundary invariants.
- Direct Commands:
  ```bash
  # Run mutant on a specific method or subject
  bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell.valid_coordinates?"
  bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell#to_i"

  # Run mutant on all 43 pure subjects (~50s)
  bundle exec rake mutant:pure
  ```

### 2. Type Safety Agent (`type-safety`)
- Target Quality Attribute: Type Safety (Static & Dynamic)
- Mission: Resolve static type errors reported by Steep and dynamic type contracts via RBS::Test, keeping RBS signatures in `sig/` and inline annotations synchronized with implementation code in `lib/`.
- Direct Commands:
  ```bash
  # Check static typing with Steep
  bundle exec rake typecheck

  # Regenerate signatures from inline annotations
  bundle exec rake sig

  # Verify runtime type contracts dynamically
  TEST_WORKERS=1 bundle exec rake test:rbs
  ```

### 3. Specification Compliance Agent (`spec-compliance`)
- Target Quality Attribute: Specification Conformance / Interoperability
- Mission: Ensure structural compliance with ECMA-376 Part 4 XML Schema definitions (SpreadsheetML, DrawingML, ChartML, DocProps) and Microsoft .NET Open XML SDK compatibility.
- Direct Commands:
  ```bash
  # Run ECMA-376 XSD schema validation suite
  bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb

  # Run Microsoft .NET Open XML SDK E2E interoperability suite (opt-in)
  bundle exec rake test:e2e
  ```

### 4. Memory Stability Agent (`memory-stability`)
- Target Quality Attribute: Resource Efficiency / Memory Stability ($O(1)$ Scalability)
- Mission: Enforce the library invariant of $O(1)$ constant streaming memory profile (retained memory < 5MB across 10,000+ rows) and eliminate unnecessary intermediate object allocations.
- Direct Commands:
  ```bash
  # Run performance and memory test suite
  bundle exec rake test:perf

  # Run ecosystem benchmark (throughput and allocation comparison)
  ruby -Ilib benchmark.rb
  ```

### 5. Throughput & Complexity Agent (`throughput-performance`)
- Target Quality Attribute: Processing Speed, Allocation Efficiency & Linear Complexity ($O(N)$)
- Mission: Guard against quadratic computational regressions ($O(N^2)$) in row and cell processing loops, bound deterministic object allocations per row, and monitor iterations-per-second (IPS) on critical hot paths.
- Scope: Parser scanning loops, XML builder escaping, column letter conversions, number formatters, large dataset scaling.
- Direct Commands:
  ```bash
  # Run throughput allocation bounds and scaling ratio test
  bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb

  # Run core hot-path microbenchmarks (benchmark-ips)
  bundle exec rake benchmark:ips
  ```

### 6. Security & Resilience Agent (`security-resilience`)
- Target Quality Attribute: Availability, Security & Attack Resilience
- Mission: Ensure complete resistance to Regular Expression DoS (ReDoS), decompression bombs (ZIP bombs), XML entity expansion (Billion Laughs), CFB circular loops, formula injection, and vulnerable dependencies.
- Scope: Regex patterns across parsers and formatters, compression stream limits, cryptographic operations, dependency advisories.
- Direct Commands:
  ```bash
  # Run ReDoS audit and formula injection mitigation test suite
  bundle exec ruby -Ilib -Itest test/security_resilience_test.rb

  # Run DoS, ZIP bomb, and encryption security audit suites
  bundle exec ruby -Ilib -Itest test/dos_protection_test.rb test/zip_bomb_test.rb test/encryption_security_audit_test.rb

  # Run dependency vulnerability check
  bundle exec rake bundle:audit
  ```

### 7. API Symmetry & Contract Agent (`api-contract`)
- Target Quality Attribute: API Symmetry & Lossless Round-Trip Fidelity
- Mission: Prevent feature and signature divergence between In-Memory (`WorksheetBuilder`, `Elements::Worksheet`) and Streaming (`StreamWriter`, `StreamSheet`) architectures, and ensure all features round-trip losslessly (`Read(Write(X)) = X`).
- Scope: Public method reflections, parameter signatures, reader query interfaces, property-based fuzzing.
- Direct Commands:
  ```bash
  # Run method reflection and parameter symmetry test
  bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb

  # Run contract test suite
  bundle exec rake test:contract

  # Run property-based testing (PBT) suite
  bundle exec rake test:pbt
  ```

### 8. Concurrency & Ractor Safety Agent (`concurrency-safety`)
- Target Quality Attribute: Concurrency Safety & Ractor Isolation
- Mission: Guarantee zero global mutable state, 100% constant shareability across the `Xlsxrb` namespace, and isolated, deterministic parallel execution across threads and Ruby 4.0+ Ractors.
- Scope: Constant freezing, Ractor shareability scanning, parallel multi-worker write/read/modify pipelines.
- Direct Commands:
  ```bash
  # Run thread and Ractor concurrency test suite
  bundle exec ruby -Ilib -Itest test/concurrent_test.rb
  ```

### 9. Visual Fidelity & Wasm Compatibility Agent (`visual-fidelity`)
- Target Quality Attribute: Visual Rendering Fidelity & WebAssembly Compatibility
- Mission: Prevent visual regressions (color loss, clipped borders, misaligned charts) when opening spreadsheets in office engines (LibreOffice Calc), and verify browser WebAssembly playground execution and headers.
- Scope: Headless LibreOffice pixel comparisons, documentation server headers (COOP/COEP), playground preset scripts.
- Direct Commands:
  ```bash
  # Run headless LibreOffice visual regression testing (VRT)
  bundle exec rake test:visual

  # Run documentation server and Wasm playground validation
  bundle exec rake doc:check

  # Regenerate golden visual baselines after intentional rendering updates
  bundle exec rake visual:baseline
  ```

---

## 3. Pre-Push Verification

The standard pre-push hook runs the complete verification suite without requiring any custom setup:

```bash
# Standard pre-push checks
bin/pre-push

# Complete pre-push checks (including mutation and RBS runtime validation)
bin/pre-push --all
```

---

## 4. Instructions for Generative AI & Autonomous Coding Agents

When any LLM-powered coding assistant or autonomous agent works on this repository, adopt the specialized QA role corresponding to the modified components:

1. Modifying pure algorithms, converters, or boundary validators:
   - Adopt the Test Rigor role (`test-rigor`).
   - Run `bundle exec rake mutant:pure`.
   - Never accept surviving mutants; preserve the 100.00% kill rate.

2. Modifying method signatures, interfaces, or public APIs:
   - Adopt the Type Safety role (`type-safety`).
   - Keep inline RBS annotations (`#: (...) -> ...`) updated in `lib/`.
   - Regenerate signatures using `bundle exec rake sig` and verify zero errors with `bundle exec rake typecheck`.

3. Modifying XML generation, serialization, or writer logic:
   - Adopt the Specification Compliance role (`spec-compliance`).
   - Run `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb` to verify ECMA-376 Part 4 XSD compliance.
   - Respect schema sequencing (`xs:sequence`) and required OOXML attributes.

4. Modifying row streaming, parsing loops, or memory buffers:
   - Adopt the Memory Stability role (`memory-stability`).
   - Run `bundle exec rake test:perf` to verify that retained memory remains flat ($O(1)$) and below the 5MB threshold.

5. Modifying parsing loops, hot paths, or scaling performance:
   - Adopt the Throughput & Complexity role (`throughput-performance`).
   - Run `bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb` to verify allocation bounds and linear scaling ($O(N)$).
   - Run `bundle exec rake benchmark:ips` to evaluate hot-path iterations per second.

6. Modifying regular expressions, compression, or security boundaries:
   - Adopt the Security & Resilience role (`security-resilience`).
   - Run `bundle exec ruby -Ilib -Itest test/security_resilience_test.rb test/dos_protection_test.rb test/zip_bomb_test.rb test/encryption_security_audit_test.rb`.
   - Ensure regular expressions complete non-matching inputs within milliseconds without catastrophic backtracking.

7. Adding or modifying worksheet authoring or reader APIs:
   - Adopt the API Symmetry & Contract role (`api-contract`).
   - Run `bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb` to confirm symmetry between `WorksheetBuilder` and `StreamWriter`.
   - Run `bundle exec rake test:contract` and `bundle exec rake test:pbt`.

8. Adding constants, class state, or concurrency logic:
   - Adopt the Concurrency Safety role (`concurrency-safety`).
   - Run `bundle exec ruby -Ilib -Itest test/concurrent_test.rb` to confirm that all constants remain `Ractor.shareable?` and parallel pipelines execute deterministically.

9. Modifying drawing ML, chart formatting, styles, or documentation playground:
   - Adopt the Visual Fidelity & Wasm Compatibility role (`visual-fidelity`).
   - Run `bundle exec rake test:visual` to ensure pixel fidelity with LibreOffice Calc.
   - Run `bundle exec rake doc:check` to ensure all playground presets execute cleanly.

### Detailed Role Playbooks
Reference playbooks detailing analysis techniques and workflows are available at:
- Test Rigor: [`.agents/skills/test-rigor/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/test-rigor/SKILL.md)
- Type Safety: [`.agents/skills/type-safety/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/type-safety/SKILL.md)
- Specification Compliance: [`.agents/skills/spec-compliance/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/spec-compliance/SKILL.md)
- Memory Stability: [`.agents/skills/memory-stability/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/memory-stability/SKILL.md)
- Throughput & Complexity: [`.agents/skills/throughput-performance/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/throughput-performance/SKILL.md)
- Security & Resilience: [`.agents/skills/security-resilience/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/security-resilience/SKILL.md)
- API Symmetry & Contract: [`.agents/skills/api-contract/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/api-contract/SKILL.md)
- Concurrency & Ractor Safety: [`.agents/skills/concurrency-safety/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/concurrency-safety/SKILL.md)
- Visual Fidelity & Wasm: [`.agents/skills/visual-fidelity/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/visual-fidelity/SKILL.md)
- Quality & Language Policies: [`.agents/rules/qa_agents.md`](file:///workspaces/xlsxrb/.agents/rules/qa_agents.md)
