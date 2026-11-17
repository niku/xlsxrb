# ADR 0012: The Argument Fixes What a String Means

* Status: Accepted
* Date: 2026-11-17
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb`](../../lib/xlsxrb.rb), [`Xlsxrb::Ooxml::ZipReader`](../../lib/xlsxrb/ooxml/zip_reader.rb), [`Xlsxrb::WorksheetBuilder`](../../lib/xlsxrb/worksheet_builder.rb), [`Xlsxrb::StreamWriter`](../../lib/xlsxrb/stream_writer.rb)

---

## Context

### What this ADR is for

xlsxrb is used in programs of many kinds, from scripts that read the author's own files to services that handle input from people they do not trust: uploaded workbooks, file names, text for cells. xlsxrb cannot tell which kind it is in, so it is designed for the second. Such a program is safe only if its author gets every call right, and the API decides how easy that is. The aim here is an API in which the mistake that matters for security is hard to make: a caller who writes the obvious code should not create a vulnerability without seeing it, and when they pass the wrong thing, the result should be an error and not a different behavior.

A review of the library against that aim found two arguments that fall short of it.

### Where the API falls short

Two arguments take a String and decide from its content, and from what exists on disk, whether it is data or the name of a file to read:

| Argument | Read as data when | Read as a path when |
| :--- | :--- | :--- |
| The source of `Xlsxrb.load` and `Xlsxrb.read` | It starts with a ZIP or compound file signature, or contains a NUL byte | Otherwise, and a file of that name exists |
| The value of `background_image=` | It contains a NUL byte, or no file of that name exists | Otherwise |

The caller does not choose between the two; the content of the String does. A program that passes on a String it received from elsewhere therefore lets the sender choose, and nothing in the calling code shows it. Each of the following was reproduced with xlsxrb 0.2.1. In each, the program is written as its author intended, and the sender supplies a file whose whole content is a path on the server.

1. Another user's workbook is returned as the upload.

   ```ruby
   workbook = Xlsxrb.load(params[:file].read)   # meant: the uploaded xlsx data
   render json: workbook.sheets.first.rows.map { |row| row.cells.map(&:value) }
   ```

   The upload is the 27 bytes `/srv/app/exports/payroll.xlsx`. xlsxrb opens that file and the program shows its rows to the sender. Only files that are valid workbooks can be read this way.

2. The sender learns what exists on the server.

   With the same code, the error depends on the path that was sent: a workbook loads, a file that is not a workbook raises `Xlsxrb::ZipError`, a directory raises `Errno::EISDIR`, and a name that does not exist raises `Errno::ENOENT`. A program that reports "not a valid workbook" for the first error and fails differently for the others tells the sender which paths exist, although it never meant to open one.

3. Any file the server can read is written into the workbook the sender downloads.

   ```ruby
   Xlsxrb.write(report) do |workbook|
     workbook.sheet("Report") do |sheet|
       sheet.background_image = params[:logo].read   # meant: the uploaded image data
       # ...
     end
   end
   send_file report
   ```

   The upload is the text `/srv/app/config/secret.txt`. xlsxrb reads that file and stores its bytes as the background image part, `xl/media/sheet1_bg.txt`, of the workbook, which the sender then downloads and unzips. The image is not checked to be an image, so this reads any file: credentials, keys, source code.

The detection was added for convenience: one argument takes whichever the caller has. Not every program is exposed: it takes one that passes on a String from someone it does not trust. For those that do, the third case discloses arbitrary files, and in none of the three does the calling code give any sign of the mistake. That is what this ADR sets out to remove.

The arguments next to these two have one meaning each, and no content is inspected:

| Argument | A String is |
| :--- | :--- |
| The destination of `Xlsxrb.write` | A path |
| The image given to `image` | Image data |

A formula is in the same position since the detection of text starting with `=` was removed: such text is text, and a formula is written with `Xlsxrb.formula`.

### How Ruby and other libraries settle it

How others answer the same question was checked on 2026-11-17, by running the code where a version is given and from the project's documentation otherwise.

| Library | A path | Data in memory |
| :--- | :--- | :--- |
| Ruby 4.0: `File.open`, `CSV.read`, `JSON.load_file`, `YAML.load_file` | A String (or `Pathname`) | Not accepted |
| Ruby 4.0: `CSV.parse`, `JSON.parse`, `YAML.load` | Not accepted | A String |
| rubyzip 3.7.0 | `Zip::File.open(string)` | `Zip::File.open_buffer(string_or_io)` |
| rubyXL | `RubyXL::Parser.parse(path)` | `RubyXL::Parser.parse_buffer(buffer)` |
| Roo | `Roo::Spreadsheet.open(path)` | A `File` instance |
| Creek | `Creek::Book.new(path)`; a URL only with `remote: true` | Not documented |
| Nokogiri 1.19.7 | Not accepted; an open `File` is passed | `Nokogiri::XML(string)` |
| openpyxl (Python) | `load_workbook("book.xlsx")` | A file-like object, such as `BytesIO` |
| pandas (Python) | `read_excel("book.xlsx")` | A file-like object |
| excelize (Go) | `OpenFile(filename string)` | `OpenReader(r io.Reader)` |
| ExcelJS (JavaScript) | `xlsx.readFile(filename)` | `xlsx.load(data)`, `xlsx.read(stream)` |
| SheetJS (JavaScript) | `XLSX.readFile(filename)` | `XLSX.read(data)`, with a `type` option |

Three things follow.

* None of them decides between a path and data by looking at the String. The two are told apart by the name of the method, by the type of the argument, or by an explicit option. xlsxrb's detection is the exception.
* Where a library reads a file format from disk, as the spreadsheet and ZIP libraries do, a plain String is a path in every one of them. Data in memory is given as a stream or to a method of its own. A String is data only where the format is text that programs commonly hold in memory: JSON, YAML, CSV, XML.
* Where a String once meant more than one thing, the extra meaning has been taken away. In Ruby, `open` no longer fetches a URL (3.0; `URI.open` is explicit) and no longer runs a command for a String starting with `|` (deprecated in 3.3, removed in 4.0, where it now raises `Errno::ENOENT`); `YAML.load` no longer builds arbitrary objects by default (Psych 4, Ruby 3.1). pandas deprecated passing the JSON text itself to `read_json`, which takes a path, in 2.1, and asks for `StringIO`.

### What was considered

Each option was judged by what happens to a caller who passes on an untrusted String, and by what a caller who passes the wrong kind sees.

1. Keep the detection.
   No code changes. The caller's only defense is to know the detection rule and wrap data in a stream. A caller who does not know it is not protected, and those are the callers this ADR is for.

2. Keep the detection, add an explicit form, and document it.
   Nothing breaks, and a careful caller has a safe way to write the call. The obvious call keeps the detection, so only callers who read the documentation are protected. A default that is safe was preferred to a safe option, for the same reason.

3. A String is always data; a file is given as a `Pathname`.
   No String handed to xlsxrb for reading is ever opened as a file, and a path passed by mistake is rejected as not being xlsx data, with an error that can show the `Pathname` form. Reading on demand from a large file is kept, since a `Pathname` is opened the way a path String is today. Against it: `Xlsxrb.load("book.xlsx")`, the call in most existing code and the way a path is given to `File.open` and `CSV.read`, stops working; none of the libraries above requires a `Pathname` for a path; and a source is written differently from the destination of `Xlsxrb.write`. Separate methods for files, in the manner of `YAML.load_file`, would read more naturally than `Pathname` but add entry points to the four of ADR 0007.

4. A String is always a path; data is given as a stream.
   The common call is unchanged, and it is what every spreadsheet and ZIP library in the table above does. A program that passes xlsx data as a String fails on every valid workbook, so it is corrected to pass a stream, and from then on no String it received from elsewhere reaches the argument. This is also the change pandas made to `read_json`.

Options 3 and 4 both remove the mistake the caller cannot see. They do not differ in what they prevent: a program that passes an untrusted path opens an untrusted file under either, and that is the caller's to prevent, as with `File.open`. What is left to option 3 is that one sentence covers every String on the reading side, and that `Pathname(...)` marks in the code where a file is opened. Those were not judged worth changing the call that most code makes. Option 4 was chosen.

For `background_image=` the same question has the other answer. Making its String a path would match the source of `Xlsxrb.load`, but the argument beside it, the image given to `image`, takes data, and two ways of placing an image should agree. Data is also the meaning under which xlsxrb opens no file while writing.

---

## Guiding Principles

1. Leave the caller no mistake they cannot see:
   Where a wrong argument could do harm, it produces an error at the call. It does not produce another behavior that the calling code gives no sign of.
2. The argument fixes the meaning:
   What a String means is decided by which argument it is passed as. It is not decided by inspecting its content or by what exists on disk.
3. A place is a path, and content is data:
   For an argument that says where to read from or write to, a String is a path. For an argument that holds the content itself, such as an image, a String is that content.
4. The other kind has its own form:
   Data for an argument that takes a place is passed as a stream (an IO, such as `StringIO`). A file for an argument that takes content is passed as its bytes, read by the caller.
5. The wrong kind is an error that says what to pass:
   A String of the other kind is not reinterpreted. It is rejected, and the error names the form to use.
6. Safe by default:
   Where convenience and Principle 1 conflict, the default is the safe behavior. Protection that a caller has to know about and ask for does not count.

---

## Decision

* The source of `Xlsxrb.load` and `Xlsxrb.read` is a path when it is a String and a stream when it is an IO. xlsx data held in memory is passed as `StringIO.new(data)`.
* A String source that holds xlsx data instead of a path is rejected with an error that says to pass a stream. A String that names no file is reported as a missing file, as `File.open` reports it.
* The value of `background_image=` is image data when it is a String, as the image given to `image` is. A file is passed as its bytes, with `File.binread`.
* The destination of `Xlsxrb.write` and the image given to `image` are unchanged.
* Reading a background image String as a path stops in the next patch release, without a warning first. It discloses arbitrary files (case 3 above), which is the exception for security fixes in the deprecation policy in [docs/ARCHITECTURE.md](../ARCHITECTURE.md). A caller that passed a path gets an image part holding the path text instead of the image, and the `Security` section of `CHANGELOG.md` says to pass `File.binread(path)`.
* Reading a String source as data follows the ordinary path of that policy: it warns, with the form to use instead, and stops in the next minor release. What it allows is narrower (cases 1 and 2), and passing data as a String is common in existing code, so those callers are given a release in which to change.
* An argument added later follows Principles 1 to 6.

This ADR does not change ADR 0007: the four entry points stay as they are, and only what their source may be is defined here.

---

## Consequences

### Positive

* A String that a program received from elsewhere and passes on as data cannot make xlsxrb open a file: data is not passed as a String where a path is expected, and where data is expected no file is opened.
* A reader of the calling code can tell from the argument what a String is, without knowing its content.
* A mistake is reported at the call, with the correction, instead of producing a different result.
* Reading a file, the most common call, is written as before, as the spreadsheet and ZIP libraries surveyed above and the destination of `Xlsxrb.write` have it.

### Negative / Trade-offs

* One argument no longer takes whichever the caller has. Callers that pass xlsx data as a String wrap it in `StringIO`, and callers that pass a file path to `background_image=` read the file themselves.
* Callers that pass a file path to `background_image=` are broken by a patch release, with no warning before it: their workbooks get a background that is not an image until the call is changed. This is the cost of closing case 3 at once.
* Until the next minor release, a String source is still read as data, so cases 1 and 2 remain possible for programs that have not changed. The warning tells them.
* The meaning of a String differs between arguments, a path for a source and data for an image, so it has to be learned per argument. The rule in Principle 3 is what keeps this predictable.
* An untrusted String passed as a path opens the file it names. xlsxrb cannot tell such a path from an intended one; checking it stays with the caller.
* The mistake is found when the code runs, not before.
