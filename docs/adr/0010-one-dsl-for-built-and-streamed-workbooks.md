# ADR 0010: One DSL for Built and Streamed Workbooks

* Status: Accepted
* Date: 2026-11-16
* Deciders: xlsxrb maintainers
* Supersedes: [ADR 0005](0005-dsl-command-void-return-and-property-accessors.md)
* Relevant Components: [`Xlsxrb::WorkbookBuilder`](../../lib/xlsxrb/workbook_builder.rb), [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb), [`Xlsxrb::Elements`](../../lib/xlsxrb/elements.rb)

---

## Context

xlsxrb writes workbooks in two modes with one DSL: building a workbook in memory, and streaming it row by row. ADR 0005 required that code written for the in-memory builder run unchanged on the streaming writer, and that DSL commands return nothing.

Three things have changed since:

1. The two modes diverge where a sheet is created or looked up. The same lookup creates a sheet in one mode and refers to an existing sheet in the other, so code written for one mode fails in the other. Creating a sheet in memory returned the sheet while it was still being built, which differs from the sheet in the finished workbook.
2. Cells in a built workbook differ from the same cells read back from the written file, in how they report their style.
3. Users need random access in built workbooks: writing any sheet, row, or cell in any order. A streaming writer cannot do this, so code written for the builder can no longer be required to run on the streaming writer. The direction of ADR 0005's interchangeability rule has to be reversed.

This ADR restates the parts of ADR 0005 that still hold and replaces it.

---

## Guiding Principles

1. Streamed code runs when built:
   Code that writes a workbook with the streaming writer produces the same workbook when it is given to the in-memory builder. The builder also accepts writes in any order. The streaming writer rejects, with an error, any write it cannot make in order. Both modes expose the same commands with the same parameters.
2. A built workbook is what is read back:
   The workbook that building returns equals the workbook read back from the file it writes.
3. Return what the caller uses next:
   Return values follow Ruby's conventions. Appending with `<<` returns the receiver, as it does on arrays and IO. An assignment returns the assigned value. A method that makes a value or reads state returns that value. A method that opens a sheet returns the writer for that sheet. Every other command returns nothing. Whatever is returned behaves the same in both modes: content that is still being built, and definitions that could be changed after they are used, are not returned.
4. One way to read and one way to set each property:
   A property with a single value has a reader without arguments and an assignment writer. A property with several options is set by calling it with options or by assignment, and read by calling it without arguments, which returns an immutable value object.
5. Boolean queries end with `?`:
   Methods that answer a yes-or-no question about a value object, reader, or row end with `?`. Bare readers for boolean state are deprecated.

---

## Decision

### 1. Sheets

* Creating a sheet makes it the current sheet and returns the writer for it, in both modes. This is the same writer that looking the sheet up returns.
* Looking up a sheet by name or 0-based index returns the writer for that existing sheet. A lookup never creates a sheet, and looking up a sheet that does not exist is an error. A lookup without an argument returns the current sheet, and creates one only when there is none.
* In the builder, every sheet can be written until the build ends. In the streaming writer, only the current sheet can be written, until the next sheet is created.

### 2. Rows and cells by position

* Rows and cells can be written at a position as well as appended. Positions use the same forms that reading accepts: an A1 reference, or 0-based row and column indexes (ADR 0004).
* Appending writes after the last row written so far.
* Writing a row at a position that already has one replaces that row, including its row attributes. Writing a cell replaces only that cell.
* The streaming writer accepts a position that does not go back, and holds at most the row being written. A position before it is an error.
* The builder does not read cells back while building. Before the build ends, a cell does not have its final style, so a value read then would differ from the built workbook (Principle 2). Cells are read from the workbook that building returns.

### 3. Styles

* Defining a style returns nothing. A style is referred to by the name the caller gave it, and a definition that could still be changed after cells use it would style them differently in the two modes.
* The builder accepts a style name before the style is defined. The streaming writer requires the definition first, because it writes each row when it receives it.
* A row, cell, or column style uses the last definition of that style made before it was written. If no definition was made before it, it uses the first definition. A row or cell written again uses the definition made before the new write.
* A built cell reports its style as the read-back cell does (Principle 2): the number format of its resolved style, style index 0 for the default style, and a numeric style reference as a style index only.

### 4. Where details are kept

This ADR fixes the rules. The methods each mode provides, their parameters, and their return values are checked by the contract tests that compare the two modes. Version-by-version changes and deprecations are recorded in `CHANGELOG.md`. Changes that bring the library in line with these rules follow the deprecation policy in [ARCHITECTURE.md](../ARCHITECTURE.md).

---

## Consequences

### Positive

* One rule tells which code runs in which mode: anything that streams also builds, with the same result.
* Moving from streaming to building never needs code changes.
* A built workbook can be compared directly with one that is read back, so tests and callers see one model of a cell.
* Built workbooks can be written in any order.

### Negative / Trade-offs

* The builder accepts code that the streaming writer rejects. Moving from building to streaming can require writing in order. The streaming writer reports each such write as an error rather than writing a different workbook.
* Only `<<` can be chained. Other commands return nothing, so each is written as its own statement.
* Bringing earlier releases into line with these rules changes what some methods return, which breaks callers that used the old return values. A changed return value cannot always be detected at runtime, so some of these changes can be announced only in the changelog.
