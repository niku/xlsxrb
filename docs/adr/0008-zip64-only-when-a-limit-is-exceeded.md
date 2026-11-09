# ADR 0008: ZIP64 Records Only When a ZIP Limit Is Exceeded

* Status: Accepted
* Date: 2026-11-09
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb::Ooxml::ZipWriter`](../../lib/xlsxrb/ooxml/zip_writer.rb), [`Xlsxrb::Ooxml::ZipReader`](../../lib/xlsxrb/ooxml/zip_reader.rb)

---

## Context

The classic ZIP format stores entry sizes and file offsets in 32-bit fields and the entry count in 16-bit fields. ZIP64 (APPNOTE 4.3.9, 4.4.13, 4.5.3) extends them to 64 bits through an extra field in the local and central directory headers, an 8-byte data descriptor, and a ZIP64 end of central directory record with its locator.

Up to 0.1.18, `ZipWriter` packed every value as 32 or 16 bits without checking. A worksheet whose XML exceeded 4 GiB uncompressed produced a corrupt XLSX without raising an error: the sizes wrapped around, and readers that check sizes or CRCs rejected the entry.

APPNOTE 4.3.9.1 allows ZIP64 to be used regardless of size. An XLSX file, however, is an Open Packaging Conventions package, and ECMA-376 Part 2 Annex B narrows what a package producer may write:

* Table B.1 marks the ZIP64 end of central directory record (V1) and its locator as "used only when needed" on production, and B.4 defines that value as "the package implementer shall use the record only when needed to store data in the ZIP file". The V2 record is not to be used.
* Table B.6 requires producers to support the ZIP64 extended information extra field (`0x0001`), and Table B.3 the version needed to extract 4.5 ("File uses ZIP64 format extensions").
* Table B.1 makes data descriptors optional on production, and Table B.5 allows general purpose bit 3.
* B.2 requires every central directory file header to hold the same values as its local file header and data descriptor, except for the zeroed fields that bit 3 allows.

---

## Options Considered

1. Always write ZIP64 records.
   One code path, and every streamed entry can carry a ZIP64 extra field in its local header from the start. Writing the ZIP64 end of central directory record for every archive would contradict OPC Annex B ("used only when needed"). Readers also differ in how they treat ZIP64 records on small files. In the Java ecosystem, Apache Commons Compress `Zip64Mode.Always` (used by Apache POI's SXSSF) produced files that some versions of Excel, LibreOffice, and 7-Zip refused to open. Commons Compress later added `Zip64Mode.AlwaysWithCompatibility` as a workaround, reports on whether Excel accepts its output are mixed, and the "xlsx zip64 compatibility" issue (COMPRESS-474) remains open. Every workbook would carry this risk to support files above 4 GiB, which are rare.
2. Raise an error when a limit is exceeded.
   Prevents silent corruption, but leaves callers unable to write the file. Other Ruby writers, such as xlsxtream through zip_kit, already write these files.
3. Write ZIP64 records only for values that exceed a limit.
   Archives below the limits stay byte-for-byte identical to plain ZIP output. This follows OPC Annex B and is the behavior of zip_kit and Go's `archive/zip`.

---

## Decision

Adopt option 3. `ZipWriter` switches to ZIP64 per value, only when that value is at or above the limit (`0xFFFFFFFF` for sizes and offsets, `0xFFFF` for the entry count, since those values are the ZIP64 markers):

* Entries whose sizes are known before the local header is written (`add_entry`, `add_binary_entry`, `copy_raw_entry`) get a ZIP64 extra field in the local header carrying both sizes, as APPNOTE 4.5.3 requires.
* A streamed entry (`start_entry` / `write_data` / `finish_entry`) on seekable output gets its local header patched with the final sizes. If the sizes outgrow 32 bits, it instead ends with a data descriptor that stores both sizes as 8-byte values, and the patched local header sets general purpose bit 3 and version needed 4.5.
* A streamed entry on unseekable output cannot have its local header patched, and its size is unknown when that header is written. Its local header states version needed 4.5 from the start, with bit 3 set, and the entry always ends with an 8-byte data descriptor; its central directory header lists both sizes in a ZIP64 extra field. This is the one case where ZIP64 records appear below the limits, the same choice Commons Compress makes in `Zip64Mode.AsNeeded` for entries of unknown size on non-seekable output.
* Each central directory header gets a ZIP64 extra field listing only the values whose 32-bit field holds `0xFFFFFFFF`, in the fixed order uncompressed size, compressed size, local header offset.
* A ZIP64 end of central directory record and locator are written when the entry count, the central directory size, or its offset exceeds a limit.

`ZipReader` reads the ZIP64 end of central directory record in addition to the ZIP64 extra fields it already handled.

---

## Consequences

### Positive
* Workbooks below the limits are unchanged; this was verified by comparing the bytes written before and after the change.
* A 4.2 GiB streamed worksheet entry passes Info-ZIP `unzip -t`, Python's `zipfile.testzip()`, and reads back through `ZipReader`.
* XLSX files with a 4.5 GB worksheet part, written to seekable output, to unseekable output, and at compression level 0 (so that later entries and the central directory start past 4 GiB), open in the Open XML SDK 3.0.2 on .NET 8 (`SpreadsheetDocument.Open` and `OpenXmlReader`), and `System.IO.Compression.ZipFile` reports the 64-bit sizes.
* Microsoft Excel (Microsoft 365 for Mac) opens XLSX files with a 4.5 GB worksheet part written to seekable output, written to unseekable output, and written at compression level 0, as well as small workbooks written to unseekable output, without a repair prompt.
* Unit tests lower the writer's ZIP64 threshold to exercise every ZIP64 path with small archives, and cross-check them with `unzip -t`.

### Negative / Trade-offs
* Excel decides the width of a data descriptor from the local header's version needed, not from the central directory. Two other layouts for unseekable output were tried and rejected because Excel offered to repair the file: version 2.0 in the local header with 4.5 in the central directory, and version 2.0 in both (which satisfies OPC B.2 but leaves the 8-byte descriptor undeclared in the local header).
* Every workbook streamed to unseekable output (a pipe or a streaming HTTP response, for example) now carries ZIP64 records on its streamed entries, even when it is small. Readers without ZIP64 support cannot read such files. Writing to a file or a `StringIO` is unaffected.
* Streamed local headers declare ZIP64 through the version needed field and the 8-byte data descriptor only; they carry no ZIP64 extra field, because the sizes are unknown when the header is written. Readers that use the central directory (Excel, Info-ZIP, Python `zipfile`, .NET, `ZipReader`) read them correctly.

---

## References

* [PKWARE APPNOTE.TXT](https://pkware.cachefly.net/webdocs/casestudies/APPNOTE.TXT) sections 4.3.9, 4.4.13, 4.5.3
* ECMA-376 Part 2 (Open Packaging Conventions), Annex B "Constraints and clarifications on ZIP features", Tables B.1 to B.6
* [Commons Compress `Zip64Mode`](https://commons.apache.org/proper/commons-compress/jacoco/org.apache.commons.compress.archivers.zip/Zip64Mode.java.html)
* [COMPRESS-474: xlsx zip64 compatibility](https://issues-test.apache.org/jira/browse/COMPRESS-474)
* [Apache POI users: whether SXSSFWorkbook should default to a different Zip64Mode](https://go.mail-archive.com/user@poi.apache.org/msg14186.html)
