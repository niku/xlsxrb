# ADR 0011: Count Excel's Text Limits as Excel Does

* Status: Accepted
* Date: 2026-11-16
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::WorkbookBuilder`](../../lib/xlsxrb/workbook_builder.rb), [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Elements::Worksheet`](../../lib/xlsxrb/elements/worksheet.rb)

---

## Context

Excel limits the length of some text: a cell holds at most 32,767 characters, and a sheet name at most 31. In strict Excel mode, xlsxrb rejects text over these limits, so that it does not write a file that Excel has to repair. It counted the characters of the text, as Ruby counts them.

"Characters" can be counted in several ways that give different numbers: bytes, Unicode characters, UTF-16 code units, and what Excel's `LEN` function reports. They differ for text with characters outside the Basic Multilingual Plane, such as emoji and some CJK Extension B ideographs, which are one Unicode character but two UTF-16 code units.

Files written with each kind of text were opened in Excel for Mac on 2026-11-16:

| Text | Unicode characters | UTF-16 code units | Excel |
| :--- | ---: | ---: | :--- |
| Cell: `x` × 32,767 | 32,767 | 32,767 | Opens |
| Cell: `x` × 32,768 | 32,768 | 32,768 | Repairs the file; keeps 32,767 |
| Cell: `あ` × 20,000 (60,000 bytes) | 20,000 | 20,000 | Opens |
| Cell: `😀` × 16,383 | 16,383 | 32,766 | Opens |
| Cell: `😀` × 16,384 | 16,384 | 32,768 | Repairs the file; keeps 16,383 emoji |
| Sheet name: `x` × 31 | 31 | 31 | Opens |
| Sheet name: `あ` × 31 | 31 | 31 | Opens |
| Sheet name: `😀` × 15 | 15 | 30 | Opens |
| Sheet name: `😀` × 16 | 16 | 32 | Repairs the file; the name is cut inside the last emoji |

Excel counts these limits in UTF-16 code units. Bytes do not matter. Excel's `LEN` reports 16,383 for 16,383 emoji, so the length Excel shows is not the length its limits count. Counting Unicode characters let strict mode write files that Excel repairs.

---

## Guiding Principles

1. Measure a limit as Excel does:
   A length limit of Excel is counted in the unit Excel counts it in. For the text limits above, that is UTF-16 code units, not bytes, Unicode characters, or what `LEN` reports.
2. Strict mode rejects; it does not repair:
   In strict Excel mode, text over a limit is an error. xlsxrb does not cut it, so the caller decides what to keep. Where xlsxrb shortens text itself, as when it makes a sheet name from arbitrary text, it shortens to the limit in the same unit and does not split a character.
3. Check Excel, and keep the result:
   How Excel counts or enforces a limit is settled by opening files in Excel, not by documentation or by the behavior of other libraries. The result is recorded in an ADR, as in this one.

---

## Decision

* Cell text and sheet names are measured in UTF-16 code units, against 32,767 and 31.
* In strict Excel mode, both modes of writing reject text over these limits with an error that gives the length in UTF-16 code units. Outside strict mode, the text is written as given.
* Checking a sheet name for validity, and making a sheet name from arbitrary text, use the same measure.
* A text limit added later follows Principles 1 to 3.
* The contract tests check written files against these limits (`test/support/excel_rules.rb`), so a writer that breaks them is found without Excel.

---

## Consequences

### Positive

* Strict mode no longer writes files that Excel repairs because of text with emoji or other characters outside the Basic Multilingual Plane.
* The rule is the same for every text limit, and its evidence is recorded.

### Negative / Trade-offs

* Text with characters outside the Basic Multilingual Plane is rejected at fewer Unicode characters than the stated limit: 16,384 emoji are too many for a cell. Callers see the length in UTF-16 code units in the error.
* Measuring in UTF-16 code units costs more than counting characters. The cost is paid only for text that may be over a limit, such as cell text longer than 32,767 bytes.
