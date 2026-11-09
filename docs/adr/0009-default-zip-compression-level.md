# ADR 0009: Default ZIP Compression Level 1

* Status: Accepted
* Date: 2026-11-09
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb`](../../lib/xlsxrb.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Ooxml::WorkbookWriter`](../../lib/xlsxrb/ooxml/workbook_writer.rb), [`Xlsxrb::Ooxml::ZipWriter`](../../lib/xlsxrb/ooxml/zip_writer.rb)

---

## Context

`Xlsxrb.write`, `StreamWriter`, `WorkbookWriter`, and `ZipWriter` accept `compression_level:` (0 to 9) and defaulted to `Zlib::DEFAULT_COMPRESSION`, which zlib maps to level 6. Exports of large sheets spend a noticeable share of their time in deflate, and the level is a trade between that time and the size of the file that is stored, uploaded, and downloaded. Reading speed does not depend on the level.

The peer writers all keep zlib's default: zip_kit (used by xlsxtream) and rubyzip (used by caxlsx, write_xlsx, and rubyXL) default to `Zlib::DEFAULT_COMPRESSION`, and libxlsxwriter (used by fast_excel) to `Z_DEFAULT_COMPRESSION`.

zlib uses two different algorithms. Levels 1 to 3 use `deflate_fast`, which takes the first acceptable match without lazy evaluation. Levels 4 to 9 use `deflate_slow`, which checks whether the next position gives a longer match. Within each group the level sets how many hash chain candidates are searched (`max_chain`: 4, 8, and 32 for levels 1 to 3; 128 for level 6) and when the search stops early (`nice_length`).

---

## Measurements

Each dataset is 100,000 rows by 10 columns (1,000,000 cells), written with `Xlsxrb.write` to a file. Every (dataset, level) pair was measured five times, with the order of levels alternating between runs to spread machine noise, and the median is reported. Read time is a streaming `Xlsxrb.read` over every cell. Environment: Ruby 4.0.9 on aarch64-linux (Docker on Apple silicon), zlib 1.3.1.

Datasets:
* Mixed: the rows of `benchmark.rb` (integers, floats, booleans, and short repeated strings).
* Text: every cell a mostly unique phrase of five lorem ipsum words and a number (about 35 bytes), written once with shared strings and once with inline strings.
* Numbers: integers and floats with two decimal places.

Write time and file size relative to level 6:

| Level | Mixed | Text, shared strings | Text, inline strings | Numbers |
| ---: | :--- | :--- | :--- | :--- |
| 0 | -28% / +802% | -14% / +1300% | -16% / +1193% | -49% / +450% |
| 1 | -26% / +41% | -17% / +5% | -17% / +15% | -43% / +17% |
| 2 | -28% / +11% | -13% / +4% | -18% / +15% | -39% / +17% |
| 3 | -22% / +7% | -7% / +11% | -7% / +20% | -29% / +13% |
| 4 | -24% / +5% | -10% / +23% | -12% / +35% | -38% / +6% |
| 5 | -18% / +2% | -10% / +8% | -7% / +3% | -27% / +2% |
| 6 | 0.81 s / 4.0 MB | 3.70 s / 6.6 MB | 2.40 s / 6.7 MB | 1.07 s / 5.9 MB |
| 9 | +264% / -5% | +125% / -5% | +66% / -2% | +327% / -7% |

Median read time showed no trend by level, varying within 7% per dataset (3.09 to 3.20 s for Mixed, 4.12 to 4.41 s and 5.80 to 6.05 s for Text, 3.07 to 3.28 s for Numbers).

Deflate alone on the worksheet XML of Mixed and Numbers (about 33 MB each), best of five, measured separately on the macOS host:

| Level | Mixed | Numbers |
| ---: | :--- | :--- |
| 1 | 0.077 s / 5.42 MB | 0.102 s / 6.90 MB |
| 2 | 0.064 s / 4.18 MB | 0.136 s / 6.91 MB |
| 3 | 0.127 s / 4.05 MB | 0.249 s / 6.69 MB |
| 6 | 0.332 s / 3.76 MB | 0.641 s / 5.92 MB |

Observations:
* Level 0 is barely faster than level 1 and produces files 5 to 14 times larger.
* Levels 1 and 2 are the fastest compressed levels; their difference in total write time is within measurement noise. Level 2 searches twice as many candidates. On repetitive XML (Mixed) it finds much longer matches, which makes it both faster and 23% smaller than level 1. On numeric XML the longer search finds nothing better, so it costs about 30% more deflate time for the same size.
* Level 4 is not a reliable middle ground: on text it produced larger files than level 1, because `deflate_slow` at level 4 uses short lazy and chain limits.
* Level 9 costs 1.7 to 4.3 times the write time of level 6 for 2 to 7% smaller files.
* The level matters most for sheets with little per-cell work. Deflate took 0.33 s of the 0.81 s Mixed write and 0.64 s of the 1.07 s Numbers write at level 6, and about 0.1 s of each at level 1. For Text, building the XML dominates.

---

## Decision

Introduce `Xlsxrb::DEFAULT_COMPRESSION_LEVEL = 1` and use it wherever `compression_level` is omitted or `nil`: `Xlsxrb.write`, `StreamWriter`, `Elements::Workbook#save` and `#to_data` (through `Xlsxrb.write`), `Ooxml::WorkbookWriter`, and `Ooxml::ZipWriter`. Explicit levels 0 to 9 and `Zlib::DEFAULT_COMPRESSION` keep working.

Level 1 was chosen over level 2 because its cost does not depend on the shape of the data: it does the least searching of any compressed level, so the write time stays lowest across datasets, and it corresponds to zlib's `BEST_SPEED`. Level 2 would give smaller files on highly repetitive sheets.

---

## Consequences

### Positive
* Writes that do not set `compression_level` are 17 to 43% faster in the measurements above.
* One constant documents the default instead of `Zlib::DEFAULT_COMPRESSION` repeated across four writers.

### Negative / Trade-offs
* Files are 5 to 41% larger than at level 6 (typically around 15 to 20%), and larger than the output of peer libraries, which use level 6. Callers that store or transfer many files can pass `compression_level: 6` (or `Zlib::DEFAULT_COMPRESSION`) to get the previous output.
* Output bytes change for callers that relied on the previous default, for example in snapshot tests.
