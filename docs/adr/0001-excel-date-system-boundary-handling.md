# ADR 0001: Excel Date System Boundary Handling and Lotus 1-2-3 Compatibility

* Status: Accepted
* Date: 2026-10-24
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::Ooxml::Utils`](../../lib/xlsxrb/ooxml/utils.rb), [`Xlsxrb::Utils`](../../lib/xlsxrb/utils.rb), [`Xlsxrb`](../../lib/xlsxrb.rb)

---

## Context

Microsoft Excel and OpenXML (ECMA-376 Part 1 §18.17.4) represent dates and times as floating-point serial numbers relative to an epoch. Two date systems are supported:
1. 1900 date system (default for Windows Excel): Day 1 is January 1, 1900 (`1900-01-01`).
2. 1904 date system (historically default for Macintosh Excel): Day 0 is January 1, 1904 (`1904-01-01`).

In the 1900 date system, Excel intentionally reproduces a bug inherited from Lotus 1-2-3: it treats the year 1900 as a leap year. Consequently, serial number `60` corresponds to February 29, 1900 (`1900-02-29`), an imaginary date that does not exist in the Gregorian calendar. All chronological dates on or after March 1, 1900 (serial number >= 61) are offset by +1 relative to elapsed calendar days since January 1, 1900.

This historical design introduces three distinct boundary edge cases when interfacing with Ruby standard library types (`Date` and `Time`), which strictly enforce the Gregorian calendar:

1. Phantom Leap Day (Serial 60): Ruby's `Date.new(1900, 2, 29)` raises `Date::Error` (invalid date).
2. Time-Only Cells (Day 0, Serial 0.0 <= x < 1.0): In Excel, cells storing time without a date (e.g. `0.5` for 12:00:00) represent elapsed fractions of Day 0 (`1900-01-00` or `1899-12-31`).
3. Pre-Epoch Dates and Negative Serials: Excel's 1900 date system does not support negative serial numbers; cells containing values < 0 display as `#####` and may trigger file corruption prompts upon opening.

---

## Ecosystem Survey

Implementations across peer libraries and language ecosystems handle these boundaries in different ways, verified by source audit and runtime boundary testing:

| Library / Platform | Base Epoch | Serial 60 (Excel Leap Day) | Serial 61 (1900-03-01) | Time-Only Serial 0.5 | Jan-Feb 1900 (1-59) | Dates Prior to Epoch / Negative Serials |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| ECMA-376 / Excel | 1900-01-01 (1) | 1900-02-29 | 1900-03-01 | 1900-01-00 / 12:00:00 | 1900-01-01 - 1900-02-28 | Displays `#####` (invalid serial) |
| .NET (`FromOADate`) | 1899-12-30 | 1900-02-28 | 1900-03-01 | 1899-12-30 12:00:00 | Off by -1 day (1899-12-31-) | Supported (negative OADate) |
| Python: `openpyxl` | 1899-12-30 | 1900-02-28 (low-level fallback) | 1900-03-01 | 1899-12-31 12:00:00 | Correct (`value += 1` if `0 < v < 60`) | Cell type validation / error |
| Java: `Apache POI` | 1899-12-30 | 1900-03-01 (dayAdjust = 0) | 1900-03-01 | 1899-12-31 12:00:00 | Correct (`startYear = 1900`) | Guarded (`isValidExcelDate = false`) |
| Ruby: `write_xlsx` | 1900-01-01 | Serial 60.0 generated | Serial 61 generated | Serial 0.5 generated | Correct (`1900-01-01` -> 1) | Guarded (`return nil if year < epoch`) |
| Ruby: `rubyXL` | 1899-12-30 | 1900-03-01 (`num < 61`) | 1900-03-01 | 1899-12-31 12:00:00 | Read is correct / Write is off by +1 | Unchecked negative float |
| Ruby: `roo` / `creek` | 1899-12-30 | 1900-02-28 | 1900-03-01 | 1899-12-30 12:00:00 | Off by -1 day (no adjustment) | Unchecked negative date |
| Ruby: `caxlsx` / `xlsxtream` | 1899-12-30 | - | - | - | Write is off by +1 (`1900-01-01` -> 2)| Unchecked negative serial |
| Ruby: `xsv` | 1899-12-30 | 1900-02-28 | 1900-03-01 | `"12:00"` (date stripped) | Off by -1 day (no adjustment) | Unchecked negative date |
| Ruby: `xlsxrb` (prior) | 1899-12-31 | 1900-03-01 | 1900-03-01 | 1899-12-31 12:00:00 | Read & write correct | Unguarded (yielded negative serials) |

---

## Guiding Principles

1. Strict ECMA-376 / Excel Alignment Where Feasible:
   Whenever a rule is defined by ECMA-376 and Microsoft Excel, and is expressible in Ruby standard types, `xlsxrb` adheres strictly to the specification. This includes exact 1:1 serial matching across 1900-01-01 to 1900-02-28 (unlike Roo/Creek/Caxlsx) and preventing generation of invalid negative date serials.
2. Pragmatic Fallback for Ruby Constraints:
   When Ruby standard library constraints (`Date` and `Time` Gregorian calendar enforcement) prevent exact representation of non-existent dates (such as February 29, 1900), `xlsxrb` adopts the most pragmatic, crash-resilient strategy validated by enterprise libraries (Apache POI and rubyXL).

---

## Decision

### 1. Serial 60 (1900-02-29) Mapping
* In [`serial_to_date`](../../lib/xlsxrb/ooxml/utils.rb), serial values `> 60` decrement by 1 before adding to `EPOCH_1900` (`1899-12-31`).
* Serial `60` and serial `61` both resolve to `Date.new(1900, 3, 1)`.
* This follows the strategy of `Apache POI` (`DateUtil`) and `rubyXL`, avoiding unexpected exceptions when parsing third-party spreadsheets containing serial 60 while preserving chronological continuity.

### 2. Time-Only Serial (Day 0) Mapping
* Serial numbers in the range [0.0, 1.0) are converted to `Time.utc(1899, 12, 31, hours, minutes, seconds)`.
* This preserves arithmetic continuity with Excel's Day 0 (`1900-01-00` / `1899-12-31`), consistent with `rubyXL` and `Apache POI`.
* Round-tripping time-only serials is supported: `datetime_to_serial(Time.utc(1899, 12, 31, 12, 0, 0))` returns `0.5`.

### 3. Pre-Epoch and Out-of-Range Guards
* Writing (`date_to_serial` / `datetime_to_serial`):
  * For 1900 date system (`date1904: false`):
    * Dates prior to `EPOCH_1900` (`1899-12-31`) raise `ArgumentError, "Date is before the 1900 date system epoch (1899-12-31)"`.
    * Dates exceeding `MAX_DATE` (`9999-12-31`) raise `ArgumentError, "Date exceeds the maximum supported date (9999-12-31)"`.
  * For 1904 date system (`date1904: true`):
    * Dates prior to `EPOCH_1904` (`1904-01-01`) raise `ArgumentError, "Date is before the 1904 date system epoch (1904-01-01)"`.
    * Dates exceeding `MAX_DATE` (`9999-12-31`) raise `ArgumentError, "Date exceeds the maximum supported date (9999-12-31)"`.
* Reading (`serial_to_date` / `serial_to_datetime`):
  * Serial values with NaN raise `ArgumentError, "Excel serial date cannot be NaN"`.
  * Serial values < 0 raise `ArgumentError, "Excel serial date must be non-negative"`.
  * Serial values exceeding `MAX_SERIAL_1900 + 1` (`2,958,466` for 1900 system) or `MAX_SERIAL_1904 + 1` (`2,957,004` for 1904 system) raise `ArgumentError, "Excel serial date exceeds maximum supported date (9999-12-31)"`.

### 4. Excel Date Validation Helper
* Add `Xlsxrb::Ooxml::Utils.valid_excel_date?(serial, date1904: false)` (delegated via `Xlsxrb::Utils.valid_excel_date?` and `Xlsxrb.valid_excel_date?`).
* Returns `true` if `serial` is a `Numeric` within [0, max_serial + 1), and `false` otherwise, mirroring Apache POI's `DateUtil.isValidExcelDate`.

---

## Consequences

### Positive
* Workbook Corruption Prevention: Rejects negative serial numbers and pre-epoch dates that corrupt Excel workbooks.
* Specification Fidelity: Maintains 100% accuracy for all genuine historical dates in 1900, avoiding the -1 day discrepancy present in Roo and Creek.
* Lossless Time Round-Tripping: Permits `1899-12-31` (Day 0) so that time-only serials (0.0 <= x < 1.0) round-trip losslessly between `Time` and `serial`.
* Diagnostic Clarity: Applications can pre-validate cell values with `valid_excel_date?` prior to conversion.

### Negative / Trade-offs
* Pre-1900 Dates as Text: Workbooks needing to store historical dates prior to 1900 must store them as formatted text strings (`@`), as required by OpenXML SpreadsheetML 1900 date system rules.
* Serial 60 One-Way Mapping: `serial_to_date(60)` maps to `1900-03-01`, but `date_to_serial(Date.new(1900, 3, 1))` yields `61`. This is an unavoidable consequence of Gregorian calendar adherence.
