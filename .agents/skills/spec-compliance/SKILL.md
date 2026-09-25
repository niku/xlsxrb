---
name: spec-compliance
description: Autonomous specification compliance agent to fix XML schema (XSD) and Microsoft Open XML SDK E2E violations based on ECMA-376 specifications.
---

# Specification Compliance Agent

The Specification Compliance Agent is responsible for upholding Specification Conformance and Real-World Interoperability, ensuring that generated XLSX packages (SpreadsheetML, DrawingML, ChartML, DocProps) comply with ECMA-376 Part 4 XML schemas and maintain compatibility with the Microsoft .NET Open XML SDK.

## Target Quality Attribute
- Specification Conformance / Interoperability: Eliminates structural schema defects, corrupt package warnings ("We found a problem with some content in it"), and third-party SDK deserialization failures.

## Responsibilities
- Diagnose schema errors from `test/xsd_validation_test.rb` (Nokogiri XSD validation) and `rake test:e2e` (.NET Open XML SDK).
- Detect sequence violations (`xs:sequence`), missing mandatory attributes, or unauthorized elements.
- Cross-reference with `docs/SPEC_SOURCES.md` and normative schemas to adjust writer implementations (`XmlBuilder`, `WorksheetWriter`, `StylesXml`).

## Native Commands

```bash
# Run ECMA-376 XSD schema validation tests (fast)
bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb

# Run Microsoft .NET Open XML SDK E2E tests (opt-in)
bundle exec rake test:e2e
```

## Schema Violation Patterns & Fixes

1. `xs:sequence` Tag Ordering Constraints:
   - OOXML schemas enforce sibling element ordering within tags like `<sheetData>`, `<cols>`, `<row>`, and `<sheetView>`.
   - Solution: Adjust writer tag output order to match the exact sequence defined in the corresponding `.xsd`.
2. Missing Required Attributes:
   - Attributes like cell references (`r="A1"`), sheet identifiers (`sheetId="1"`), or relationship IDs (`r:id="rId1"`).
3. Namespace Prefix Inconsistencies:
   - DrawingML elements require explicit namespaces (`xdr:`, `a:`, `r:`).

## Standard Workflow
1. Run `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb`.
2. Analyze Nokogiri validation messages (e.g. `Element '...': This element is not expected.`).
3. Check `test/fixtures/xsd/` or ECMA-376 specifications for expected element ordering.
4. Update XML generation logic in `lib/xlsxrb/ooxml/`.
5. Re-run `bundle exec ruby -Ilib -Itest test/xsd_validation_test.rb` to confirm all assertions pass.
