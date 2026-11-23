# ADR 0014: Build XML Through `Ooxml::XmlBuilder`

- Status: Accepted
- Date: 2026-11-23

## Context

`Ooxml::Writer` and its mixins previously assembled XML by interpolating values into string literals such as `<tag attr="#{xml_escape(val)}">`. Hundreds of attribute values and several text nodes were omitted from `xml_escape` calls until an audit found them, and several attributes were escaped twice because a helper escaped a value into a variable that was escaped again when interpolated into the enclosing tag. A static test (`test_writers_escape_every_attribute_value`) was added to inspect every attribute interpolation and concatenation across the codebase.

Cell-loop and shared-string serialization paths (`Ooxml::WorksheetWriter` and `Ooxml::WorkbookWriter#write_shared_strings_xml`) run once per cell or string and are constrained by Quality Gates 4 and 5 (under 25 objects allocated per row written, O(1) memory, linear scaling). All other XML generation runs once per workbook, worksheet part, chart, drawing, table, or property block and is not bound to row-rate allocation limits.

## Options Considered

### Option A: Rely on static source inspection alone

Keep string interpolation across all writers and rely on regex matching (`ATTRIBUTE_INTERPOLATION`, `ATTRIBUTE_CONCATENATION`) plus the manual `UNESCAPED_ATTRIBUTE_VALUES` allowlist.

Rejected because:
- Source matching only recognizes specific interpolation and concatenation forms and misses alternative string construction or double-escaping when an already escaped string is interpolated again.
- Maintaining `UNESCAPED_ATTRIBUTE_VALUES` across every writer mixes structural metadata writers with hot paths and relies on author discipline for every new XML tag.

### Option B: Commit a dynamic literal-replacement audit script and rake task

Replace literals across examples and tests with an attack string (`ATTRIBUTE_BREAKOUT`) via Ripper and instrument ZIP writers in parallel to detect unescaped XML injections.

Rejected because:
- Coverage is limited to options exercised by examples and tests (an earlier audit missed Hash cell `:ref` and `:type`, which static inspection caught).
- Tests that intentionally pass raw XML parts or binary payloads trigger false positives, requiring an ongoing allowlist and parallel task maintenance outside the pre-push hook.
- The literal-replacement audit remains useful as a manual pre-release check without adding repository surface.

### Option C: Serialize XML through `Ooxml::XmlBuilder` with guarded hot-path exclusions

Route all XML serialization outside the row and shared-string hot paths through `Ooxml::XmlBuilder`, which escapes every attribute value and text node in one place (`write_attrs`, `text`, `text_tag`, `xstring_text`), while keeping direct string building only in `WorksheetWriter` and `WorkbookWriter#write_shared_strings_xml`.

Selected because it removes string-concatenated attribute construction from ~1,000 non-hot-path sites, fixes double-escaping by passing unescaped domain values directly to `XmlBuilder`, and preserves hot-path throughput and allocation budgets.

## Principles and Rules

1. Writers outside the row-rate hot paths (`Ooxml::Writer`, its `DrawingXml`, `FeaturesXml`, and `StylesXml` mixins, and `Ooxml::Crypto::Agile`) must construct all XML elements, attributes, and text through `Ooxml::XmlBuilder` and must not contain raw XML start-tag string literals.
2. Callers of `Ooxml::XmlBuilder` must pass raw (unescaped) values so escaping happens once inside `XmlBuilder`.
3. Only `Ooxml::WorksheetWriter` and `Ooxml::WorkbookWriter` may use string concatenation for XML attributes or tags.
4. Any attribute expression in those hot-path writers must pass static inspection (`test_writers_escape_every_attribute_value`); any unescaped expression listed in `UNESCAPED_ATTRIBUTE_VALUES` must be backed by a test passing an attack string containing `"` and `<`.

## Consequences

- `Ooxml::Writer`, its mixins, and `Ooxml::Crypto::Agile` no longer call string-escaping helpers directly or appear in `UNESCAPED_ATTRIBUTE_VALUES`.
- Security tests verify both the absence of raw XML start-tag literals in `XmlBuilder`-backed writers and the single-escaping and breakout resistance of hot-path writers.
- Any future XML part writer automatically follows the `XmlBuilder` rule unless explicitly scoped and justified as a row-rate hot path.
