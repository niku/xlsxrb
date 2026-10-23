---
name: security-resilience
description: Autonomous security agent to prevent ReDoS, resource exhaustion DoS, formula injection, and cryptographic vulnerabilities.
---

# Security & Resilience Agent

The Security & Resilience Agent is responsible for guarding Availability, Cryptographic Integrity, and Injection Resistance, ensuring that `xlsxrb` securely processes untrusted XLSX archives and user-provided inputs without catastrophic backtracking or resource exhaustion.

## Target Quality Attribute
- Availability, Security & Attack Resilience: Protects against Regular Expression Denial of Service (ReDoS), decompression bombs (ZIP bombs), XML entity expansion attacks (Billion Laughs), CFB circular loops, excessive KDF iteration amplification, formula injection, and known third-party CVEs.

## Responsibilities
- Audit all regular expressions against catastrophic backtracking using 10,000+ character adversarial payloads (unclosed XML tags, deeply nested brackets, large repetitions).
- Enforce strict resource bounds on decompression (`max_uncompressed_size` per ZIP entry) across both in-memory and streaming read APIs.
- Guard cryptographic implementations: ensure constant-time password hash comparison (`OpenSSL.secure_compare`), CSPRNG salt/IV uniqueness (`SecureRandom`), and upper bounds on Agile PBKDF2 `spinCount` (max 10,000,000).
- Guard formula injection boundaries via `escape_formulas: true` to prevent CSV/formula injection vulnerabilities from executing formulas when handling untrusted user input.
- Scan runtime and development dependencies for known security advisories via Bundler Audit.

## Native Commands

```bash
# Run security and ReDoS audit test suite
bundle exec ruby -Ilib -Itest test/security_resilience_test.rb

# Run comprehensive DoS and encryption audit suites
bundle exec ruby -Ilib -Itest test/dos_protection_test.rb test/zip_bomb_test.rb test/encryption_security_audit_test.rb

# Scan dependencies for known CVEs
bundle exec rake bundle:audit
```

## Security Invariants & Attack Vectors

1. Regular Expression Denial of Service (ReDoS):
   - Never use nested repetitions (e.g. `(a+)+` or `([^>]*)*`) or overlapping disjunctions.
   - All parser regexes must evaluate 10,000+ character non-matching edge-case inputs within milliseconds (< 50ms).
2. ZIP Decompression Bomb Defense:
   - Decompression streams must track cumulative uncompressed bytes and raise `ArgumentError` when exceeding `max_uncompressed_size` (default 500MB).
   - Both `ZipReader.open` and streaming facade `Xlsxrb.read` must enforce this invariant.
3. XML Billion Laughs & Entity Limits:
   - XML parsing through SAX must enforce standard entity expansion caps (10,000 expansions max) via parser security configuration.
4. CFB Sector Chain Loop Defense:
   - CFB FAT and MiniFAT traversals must track visited sector indices to terminate circular references without entering infinite loops.
5. Formula Injection Defense:
   - When `escape_formulas: true` is enabled, string values starting with formula prefix characters (`=`, `+`, `-`, `@`) must be treated as textual cells and prevented from emitting formula `<f>` elements.

## Standard Workflow
1. Run `bundle exec ruby -Ilib -Itest test/security_resilience_test.rb test/dos_protection_test.rb test/zip_bomb_test.rb test/encryption_security_audit_test.rb`.
2. Run `bundle exec rake bundle:audit` to check gem advisories.
3. When adding or modifying regular expressions in `WorksheetParser`, `NumberFormatter`, `RelationshipsParser`, or writers, add corresponding adversarial payload cases to `test/security_resilience_test.rb`.
4. Confirm all assertions pass within configured latency timeouts.
