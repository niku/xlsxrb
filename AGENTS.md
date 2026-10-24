# Instructions for AI Coding Agents

Guidelines for LLM-powered coding assistants (Claude, Cursor, Copilot, Gemini, Antigravity, etc.) and autonomous agents contributing to xlsxrb.

## Core Policies

- Language Policy:
  xlsxrb is an open-source project intended for the global developer community. To ensure accessibility and collaboration across contributors, all code comments, documentation, test names, commit messages, and PR descriptions must be written in English.
- Writing Style & Tone:
  - Factual and Objective: Use restrained, engineering-focused language. Avoid promotional, hyperbolic, or vague adjectives (e.g., "robust", "comprehensive", "seamless", "enterprise-grade", "ultra-fast", "guarantee"). State actual behavior, verified metrics, and technical constraints directly.
  - Minimal Markdown Formatting: Avoid ornamental bold formatting (`**...**`) across table cells, list items, or section headers. Reserve bold styling strictly for genuine emphasis or critical callout banners.
- Minimal Surface Area:
  Do not introduce new files, scripts, tasks, or dependencies unless their necessity is explicitly justified. Use existing tools and tasks (`bundle exec rake test:unit`, `rake typecheck`, `rake mutant:pure`, `bin/pre-push`, etc.).
- Pre-Push Verification:
  Ensure the verification suite passes before pushing or completing a task: `bin/pre-push --all`. For intermediate commits or non-logic tasks, run lightweight verification (`bin/pre-push` or component rake tasks) to maintain fast feedback loops.
- Atomic Commits & Granularity:
  Each commit must correspond to exactly one logical change, feature, or bug fix. Never bundle multiple unrelated fixes, improvements, or features into a single commit. Every commit must encompass its specific implementation, unit/contract tests, and corresponding CHANGELOG.md entry while satisfying component quality gates.
- Changelog Policy:
  Maintain `CHANGELOG.md` following [Keep a Changelog](https://keepachangelog.com/) conventions.
  - When introducing user-facing changes (features, deprecations, breaking changes, notable bug fixes, performance improvements, or dependency changes), add concise entries under the `## [Unreleased]` section.
  - Group entries under standard categories: `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, or `Security`.
- Byte vs Character Indexing Consistency:
  When scanning XML or raw data with fast string indexing (`String#index` or `String#rindex`), never mix character-level indices with byte-level operations (`String#byteslice`, `String#getbyte`). Always convert the buffer to binary (`str.b` or `force_encoding(Encoding::BINARY)`) before scanning so that index positions correspond 1:1 to byte offsets. Mixing character indices with byte slicing causes terminal row truncation and silent data corruption when documents contain multibyte (e.g. CJK or emoji) characters.

## Quality Gates by Component

When modifying specific parts of the codebase, the following verification standards must be satisfied:

1. Pure algorithms and coordinate logic (`lib/xlsxrb/elements/`, `lib/xlsxrb/ooxml/utils.rb`, `lib/xlsxrb/dsl_helpers.rb`):
   - Quality Attribute: Test Suite Rigor / Detection Power
   - Verification: `bundle exec rake mutant:pure`
   - Requirement: Maintain a 100.00% mutation kill rate across all 43 subjects. Do not accept surviving mutants.

2. Public interfaces and method signatures (`lib/`, `sig/`):
   - Quality Attribute: Type Safety (Static & Dynamic)
   - Verification: `bundle exec rake typecheck` (and `rake sig` to synchronize)
   - Requirement: 0 Steep type errors. Inline annotations (`#: ...`) must match implementation.

3. XML serialization and OpenXML writers (`lib/xlsxrb/ooxml/writer*`):
   - Quality Attribute: Specification Conformance / Interoperability
   - Verification: `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb`
   - Requirement: Compliance with ECMA-376 Part 4 XML Schema sequence constraints (`xs:sequence`).

4. Streaming pipelines and row iteration (`lib/xlsxrb/stream_*`):
   - Quality Attribute: Resource Efficiency / Memory Stability ($O(1)$)
   - Verification: `bundle exec rake test:perf`
   - Requirement: Retained memory must remain constant (< 5MB across 10,000+ rows).

5. Throughput and computational complexity (`lib/xlsxrb/ooxml/worksheet_parser.rb`, `lib/xlsxrb/ooxml/xml_builder.rb`, `lib/xlsxrb/number_formatter.rb`):
   - Quality Attribute: Processing Speed / Linear Complexity ($O(N)$) & Low GC Churn
   - Verification: `bundle exec ruby -Ilib -Itest test/performance/throughput_test.rb` and `bundle exec rake benchmark:ips`
   - Requirement: Allocation caps (< 25 objects/row write, < 15 objects/row read), scaling ratio $\le 15.0\times$ on $10\times$ input growth.

6. Security, DoS defense, and regular expressions (`lib/xlsxrb/ooxml/`, `lib/xlsxrb/ooxml/crypto/`):
   - Quality Attribute: Availability / Attack Resilience
   - Verification: `bundle exec ruby -Ilib -Itest test/security_resilience_test.rb test/dos_protection_test.rb test/zip_bomb_test.rb test/encryption_security_audit_test.rb` and `bundle exec rake bundle:audit`
   - Requirement: Zero ReDoS catastrophic backtracking (< 50ms on 10k+ adversarial inputs), ZIP bomb bounds, zero CVEs.

7. Builder / Reader API consistency (`lib/xlsxrb/worksheet_builder.rb`, `lib/xlsxrb/stream_writer.rb`, `lib/xlsxrb/elements/worksheet.rb`, `lib/xlsxrb/stream_sheet.rb`):
   - Quality Attribute: API Symmetry / Lossless Round-Trip Fidelity
   - Verification: `bundle exec ruby -Ilib -Itest test/contract/api_symmetry_test.rb` and `bundle exec rake test:contract test:pbt`
   - Requirement: 0 method or arity discrepancies between builders, lossless `Read(Write(X)) = X`.

8. Concurrency, thread safety, and Ractor isolation (`lib/`):
   - Quality Attribute: Concurrency Safety / Ractor Shareability
   - Verification: `bundle exec ruby -Ilib -Itest test/concurrent_test.rb`
   - Requirement: 100% constant shareability across `Xlsxrb` namespace, 0 `Ractor::IsolationError` in parallel pipelines.

9. Visual rendering and WebAssembly playground (`examples/visual/`, `docs/visual/`, `docs/preview.html`):
   - Quality Attribute: Visual Rendering Fidelity / WebAssembly Compatibility
   - Verification: `bundle exec rake test:visual` and `bundle exec rake doc:check`
   - Requirement: Pixel difference within thresholds ($\le 100$px tables, $\le 500$px charts), 100% passing Wasm playground presets with COOP/COEP headers.

Detailed procedural playbooks are documented in `.agents/skills/` and [docs/QA_AGENTS.md](docs/QA_AGENTS.md). For local environment setup (Dev Container) and test command references, see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).
