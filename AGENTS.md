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
  Ensure the entire verification suite passes before submitting changes: `bin/pre-push --all`.
- Changelog Policy:
  Maintain `CHANGELOG.md` following [Keep a Changelog](https://keepachangelog.com/) conventions.
  - When introducing user-facing changes (features, deprecations, breaking changes, notable bug fixes, performance improvements, or dependency changes), add concise entries under the `## [Unreleased]` section.
  - Group entries under standard categories: `Added`, `Changed`, `Deprecated`, `Removed`, `Fixed`, or `Security`.
  - When cutting a release, rename `[Unreleased]` to the new version header with the release date (e.g., `## [0.1.13] - 2026-09-25`) and add a new empty `## [Unreleased]` section above it.

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

Detailed procedural playbooks are documented in `.agents/skills/` and [docs/QA_AGENTS.md](docs/QA_AGENTS.md). For local environment setup (Dev Container) and test command references, see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).
