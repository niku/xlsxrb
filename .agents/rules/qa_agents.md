# xlsxrb Autonomous Quality Assurance (QA) Guidelines

xlsxrb is an open-source project. All modifications and contributions must adhere to the following quality gates and guidelines.

## Language Policy
As an open-source library intended for global use, all code comments, documentation, commit messages, tests, and CLI outputs must be written in English unless explicitly requested otherwise by the user.

## Writing Style & Tone
- Factual and Objective: Use restrained, engineering-focused language. Avoid promotional, hyperbolic, or vague adjectives (e.g., "robust", "comprehensive", "seamless", "enterprise-grade", "ultra-fast", "guarantee"). State actual behavior, verified metrics, and technical constraints directly.
- Minimal Markdown Formatting: Avoid ornamental bold formatting (`**...**`) across table cells, list items, or section headers. Reserve bold styling strictly for genuine emphasis or critical callout banners.

## Native Tooling Principle
To keep the repository clean and maintainable:
- Do not introduce redundant bin scripts or custom Rake tasks when existing tasks (`rake mutant:pure`, `rake typecheck`, `rake test:unit`, `rake test:perf`, `bin/pre-push`) are already available.
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

## Pre-Push Verification
Before pushing changes, run:

```bash
bin/pre-push --all
```
