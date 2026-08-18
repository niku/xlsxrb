# Specification Reference Policy

This document defines the specification reference policy for `xlsxrb`.

## Canonical Specification (Normative)
The primary reference for this library is **ECMA-376 (Office Open XML File Formats)**.
Specifically, `xlsxrb` targets the **Transitional** version of the specification (Part 4) to ensure maximum compatibility with existing applications such as Microsoft Excel, LibreOffice, and Google Sheets.

Local copies of the ECMA-376 specifications are located in the `vendor/docs/` directory:
- [Part 1: Fundamentals And Markup Language Reference](file:///workspaces/xlsxrb/vendor/docs/ECMA-376-Part1/Ecma%20Office%20Open%20XML%20Part%201%20-%20Fundamentals%20And%20Markup%20Language%20Reference.pdf)
- [Part 2: Open Packaging Conventions](file:///workspaces/xlsxrb/vendor/docs/ECMA-376-Part2/ECMA-376-2_5th_edition_december_2021.pdf)
- [Part 3: Markup Compatibility and Extensibility](file:///workspaces/xlsxrb/vendor/docs/ECMA-376-Part3/Ecma%20Office%20Open%20XML%20Part%203%20-%20Markup%20Compatibility%20and%20Extensibility.pdf)
- [Part 4: Transitional Migration Features](file:///workspaces/xlsxrb/vendor/docs/ECMA-376-Part4/Ecma%20Office%20Open%20XML%20Part%204%20-%20Transitional%20Migration%20Features.pdf)

## Supplementary Specifications (Excel Real-world Behavior)
To address gaps between the official ECMA standard and actual implementations in Microsoft Excel, the following Microsoft Open Specifications are used. Note that these files are not bundled in the repository to avoid licensing/redistribution issues; instead, they are referenced via online links and versioned here:

1. **[[MS-XLSX]: Excel Extensions to OOXML SpreadsheetML Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/)**
   - **Role:** Explains Excel-specific extensions, default attributes, and schema extensions.
   - **Referenced Version:** July 2024 / Version 12.0 (or current release).
2. **[[MS-OI29500]: Office Implementation Information for ISO/IEC 29500](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/)**
   - **Role:** Identifies how Excel actually reads/writes files, including deviations and compatibility behaviors.
   - **Referenced Version:** July 2024 / Version 12.0 (or current release).
3. **[[MS-OFFCRYPTO]: Office Document Cryptography Structure](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-offcrypto/)**
   - **Role:** Provides details on encryption, passwords, and hashing algorithms used for document/sheet protection (Agile Encryption and Standard Encryption).
   - **Referenced Version:** July 2024 / Version 12.0 (or current release).
4. **[[MS-CFB]: Compound File Binary File Format](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/)**
   - **Role:** Container format (OLE structured storage) wrapping encrypted Office documents (`EncryptionInfo` and `EncryptedPackage` streams).
   - **Referenced Version:** July 2024 / Version 14.0 (or current release).

### ISO/IEC 29500 Note
ISO/IEC 29500 is contents-wise equivalent to ECMA-376. However, ISO/IEC 29500 requires paid purchase in general, whereas ECMA-376 is freely available. Therefore, we primarily cite ECMA-376 sections.

## Source Attribution Policy
When implementing features or fixing bugs that depend on specific behaviors defined in these specifications, developers should:
1. Reference the specific part, section, or page number in code comments (e.g., `# Reference: ECMA-376 Part 1, Section 18.3.1.73`).
2. Update the mapping table below to maintain a central registry of implemented standard behaviors.

## Implementation Specification Mapping Table
| Feature | Primary Standard | Specification Section | Notes |
| :--- | :--- | :--- | :--- |
| Cell Value Types | ECMA-376 Part 1 | §18.18.11 (ST_CellType) | Defines cell data types (boolean, number, string, formula, etc.) |
| Shared Strings | ECMA-376 Part 1 | §18.4 (Shared String Table) | Handling of `<sst>` and `<si>` for cell value reuse |
| Styles & Formatting | ECMA-376 Part 1 | §18.8 (Styles) | Cell style XF indexes, font, fill, border mappings |
| Hyperlinks | ECMA-376 Part 1 | §18.3.1.48 (hyperlink) | Worksheet hyperlinks referencing external URLs or internal targets |
| Document Encryption (Agile) | [MS-OFFCRYPTO] | §2.3.4 (Agile Encryption) | AES-256-CBC, PBKDF2/SHA-512, HMAC-SHA512 data integrity |
| Document Encryption (Standard) | [MS-OFFCRYPTO] | §2.3.6 (Standard Encryption) | AES-128-ECB, SHA-1 with CryptoAPI 50,000-spin key derivation |
| Encryption Container | [MS-CFB] | §2 (Compound File Structure) | Mini Stream, FAT/MiniFAT sectors, and Red-Black tree directory |
