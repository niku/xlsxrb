# Autonomous Quality Assurance (QA) Agents in xlsxrb

This document defines the architecture, responsibilities, standard commands, and autonomous workflows of the Autonomous QA Agents Squad for xlsxrb.

---

## 1. Overview & Architectural Philosophy

xlsxrb is a pure Ruby library for reading and writing Excel spreadsheets with streaming support. Rather than introducing wrapper scripts or ad-hoc tasks, the autonomous QA agents directly leverage the repository's existing verification infrastructure to guard specific quality attributes:

```
┌────────────────────────────────────────────────────────────────────────┐
│                        xlsxrb Autonomous QA Squad                      │
├──────────────────┬──────────────────┬──────────────────┬───────────────┤
│    Test Rigor    │   Type Safety    │ Spec Compliance  │MemoryStability│
│ (Mutant 100% Kill│  (Steep & RBS)   │  (ECMA-376 SDK)  │ (O(1) Memory) │
├──────────────────┼──────────────────┼──────────────────┼───────────────┤
│ bundle exec      │ bundle exec rake │ bundle exec ruby │ bundle exec   │
│ mutant run ...   │ typecheck / sig  │ test/xsd_...     │ rake test:perf│
└──────────────────┴──────────────────┴──────────────────┴───────────────┘
```

---

## 2. Specialized Agents & Quality Attributes

### 1. Test Rigor Agent (`test-rigor`)
- Target Quality Attribute: Test Suite Rigor / Detection Power
- Mission: Eliminate surviving mutants in pure logic, algorithms, and conversions to maintain a 100.00% kill rate across all 38 pure functional subjects.
- Scope: Coordinate conversions, Julian day/serial dates, binary packaging, XML escaping, row/column boundary invariants.
- Direct Commands:
  ```bash
  # Run mutant on a specific method or subject
  bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell.valid_coordinates?"
  bundle exec mutant run --usage opensource -r ./test/xlsxrb/elements_test.rb -- "Xlsxrb::Elements::Cell#to_i"

  # Run mutant on all 38 pure subjects (~28s)
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

When any LLM-powered coding assistant or autonomous agent (e.g., Claude, Cursor, Copilot, Gemini, OpenAI, or multi-agent frameworks) works on this repository, it should adopt the corresponding specialized QA role based on the quality attribute at stake:

1. When modifying pure algorithms, converters, or boundary validators:
   - Adopt the Test Rigor role.
   - Run `bundle exec mutant run --usage opensource -r ./test/xlsxrb/<test_file>.rb -- "<Subject>"` on modified methods, or run `bundle exec rake mutant:pure`.
   - Never accept surviving mutants. Add precise, boundary-conscious assertions to maintain a 100.00% kill rate.

2. When modifying method signatures, interfaces, or adding public APIs:
   - Adopt the Type Safety role.
   - Keep inline RBS annotations (`#: (...) -> ...`) updated in `lib/`.
   - Regenerate signatures using `bundle exec rake sig` and verify zero errors with `bundle exec rake typecheck`.

3. When modifying XML generation, serialization, or writer logic:
   - Adopt the Specification Compliance role.
   - Run `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb` to verify ECMA-376 Part 4 XSD compliance.
   - Respect schema sequencing (`xs:sequence`) and required OOXML attributes.

4. When modifying row streaming, parsing loops, or memory buffers:
   - Adopt the Memory Stability role.
   - Run `bundle exec rake test:perf` to verify that retained memory remains flat ($O(1)$) and below the 5MB threshold.
   - Avoid creating persistent object references inside stream callbacks.

### Detailed Role Playbooks
Reference playbooks detailing analysis techniques and workflows are available at:
- Test Rigor: [`.agents/skills/test-rigor/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/test-rigor/SKILL.md)
- Type Safety: [`.agents/skills/type-safety/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/type-safety/SKILL.md)
- Specification Compliance: [`.agents/skills/spec-compliance/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/spec-compliance/SKILL.md)
- Memory Stability: [`.agents/skills/memory-stability/SKILL.md`](file:///workspaces/xlsxrb/.agents/skills/memory-stability/SKILL.md)
- Quality & Language Policies: [`.agents/rules/qa_agents.md`](file:///workspaces/xlsxrb/.agents/rules/qa_agents.md)
