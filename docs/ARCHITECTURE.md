# Xlsxrb Architecture

Xlsxrb uses a multi-layered architecture to separate low-level OpenXML specification details from a high-level, idiomatic Ruby API. This separation of concerns ensures the library is both robust against the complex OpenXML spec and user-friendly for Ruby developers.

## Core Architecture

The library is structured into three distinct layers:

### 1. Low-Level Infrastructure (The "OOXML" Layer)

**Namespace:** `Xlsxrb::Ooxml`

**Responsibility:**
This layer directly handles ZIP extraction, XML parsing (via SAX), and XML generation. It adheres strictly to the ECMA-376 OpenXML specification.
* **`Xlsxrb::Ooxml::Parser`**: Parses `.xlsx` files and returns raw strings, hashes, and arrays representing the OpenXML structures. Supports both file paths and **IO objects**.
* **`Xlsxrb::Ooxml::Generator`**: Takes raw data and generates valid OpenXML and `.xlsx` files. Supports streaming to file paths or **IO objects**.

### 2. High-Level Domain Model (The "Elements" Layer)

**Namespace:** `Xlsxrb::Elements`

**Responsibility:**
This layer provides idiomatic, easy-to-use Ruby objects representing Excel concepts. It utilizes Ruby 3.2+ `Data` classes for immutability and precise structural definition. All domain models are encapsulated here to keep the top-level namespace clean.

**Core Objects (`Data` classes):**
* **`Xlsxrb::Elements::Workbook`**: Represents the entire file structure.
* **`Xlsxrb::Elements::Worksheet`**: Represents a single sheet.
* **`Xlsxrb::Elements::Cell`**: Represents a single cell.
* **`Xlsxrb::Elements::Row` / `Column`**: Represents formatting and properties.

**Design Principles:**
* **Zero-based Indexing:** To maintain consistency with Ruby's core language (Arrays/Enumerable), all indices (rows, columns, and worksheets) are **0-based**. For Excel-style coordination, use string references like `cell("A1")`.
* **Fail-safe Design (Lazy Validation):** Exceptions are not raised during XML parsing. Each class has an `errors` property (Array) and a `valid?` method for state verification.
* **Forward Compatibility:** All classes have an `unmapped_data` property (Hash) to ensure that any unknown XML attributes or elements are retained, preserving file integrity during round-trips.

### 3. The Facade / Entrypoint Layer

**Namespace:** `Xlsxrb` module methods

**Responsibility:**
Acts as the primary bridge, offering both In-Memory and Streaming APIs.

| Method | Type | Description |
| :--- | :--- | :--- |
| **`Xlsxrb.read(source)`** | In-Memory | Loads the entire file into a `Workbook` object. |
| **`Xlsxrb.foreach(source, **options)`** | **Streaming** | Yields each `Row` one by one. Ideal for large files. |
| **`Xlsxrb.write(target, workbook)`** | In-Memory | Saves a `Workbook` object to a file or IO. |
| **`Xlsxrb.generate(target, &block)`** | **Streaming** | Provides a DSL to stream data directly to a file/IO. |

## Benefits of this Approach

* **Rubyish Interface:** Methods like `foreach` and `generate` follow Ruby's standard library conventions (e.g., `CSV.foreach`).
* **Clean Namespace:** Users only interact with the `Xlsxrb` module. Internal models are safely isolated within `Elements`.
* **Safety & LSP Support:** `Data` objects provide clear property definitions for editor autocomplete.
* **Future-Proofing:** This structure accommodates future features like a Lazy Enumerator API without rewriting the underlying XML logic.
