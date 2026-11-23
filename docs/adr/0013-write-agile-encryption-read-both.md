# ADR 0013: Write Agile Encryption, Read Both

* Status: Accepted
* Date: 2026-11-23
* Deciders: xlsxrb maintainers
* Relevant Components: [`Xlsxrb`](../../lib/xlsxrb.rb), [`Xlsxrb::Ooxml::Crypto`](../../lib/xlsxrb/ooxml/crypto.rb), [`Xlsxrb::Ooxml::Crypto::Agile`](../../lib/xlsxrb/ooxml/crypto/agile.rb), [`Xlsxrb::Ooxml::Crypto::Standard`](../../lib/xlsxrb/ooxml/crypto/standard.rb)

---

## Context

### What this ADR is for

A caller who passes `password:` to `Xlsxrb.write` expects the file to be protected as well as the format allows. The format has two schemes, and xlsxrb 0.2 writes the weaker one unless the caller asks for the other by name. A security review of the library took up which schemes xlsxrb should write and read, so that the obvious call gives the stronger protection and the weaker one is not written by oversight.

### The two schemes

[MS-OFFCRYPTO] defines both.

| | Standard | Agile |
| :--- | :--- | :--- |
| Cipher | AES-128, ECB | AES-256, CBC |
| Key derivation | SHA-1, 50,000 iterations | SHA-512, 100,000 iterations |
| Detects a changed file | No | Yes, by an HMAC of the encrypted package |
| Written by Excel | 2007 | 2010 and later |
| `Xlsxrb.write` in 0.2 | The default | `encryption_mode: :agile` |

Under ECB, equal blocks of the package encrypt to equal blocks, and nothing in a Standard file tells a reader that the file was changed after it was encrypted.

### What was found

The following was checked from 2026-11-19 to 2026-11-23 with xlsxrb at 0.2.1 plus the fixes named below, Apache POI 5.2.5, and Excel 16.115 for Mac.

* The Agile writer wrote an HMAC that was not the one the specification describes, and the Agile reader did not check the HMAC at all. A file changed after encryption could be read without an error. Both were fixed before this ADR: the writer writes the specified HMAC, and the reader checks it and still reads files with the HMAC that earlier versions wrote.
* Excel opens the Agile and the Standard files that xlsxrb writes. It also opens an Agile file whose HMAC does not match, so it does not refuse a file for its HMAC.
* A workbook that Excel saved with a password uses Agile encryption with the parameters xlsxrb writes: AES-256 in CBC mode and SHA-512 with 100,000 iterations. Its HMAC matches under the reading of the specification that xlsxrb follows, xlsxrb reads it, and xlsxrb refuses it after a change to any one of four bytes that were tried.
* Encrypted workbooks in the test data of Apache POI and msoffcrypto-tool that Excel wrote, going by the application recorded in them, were read as well: four with Agile encryption, from Excel 2013 and Excel 2016 for Mac, with AES-128 and SHA-1 or AES-256 and SHA-512, all of which pass the HMAC check; and two with Standard encryption from Excel 2007. One of the two could not be read until a fault in reading its container was fixed, which no file written by xlsxrb or POI had shown.
* Apache POI decrypts the Agile and the Standard files that xlsxrb writes, and xlsxrb reads the files that POI encrypts with either scheme.
* The end-to-end tests with LibreOffice and the Open XML SDK write with the default, so they covered Standard only. No test read an encrypted file that another implementation had written with a known scheme.

### What was considered

1. Keep both schemes for writing, with Standard as the default.
   Nothing changes for callers. The obvious call, `password:` alone, keeps writing the weaker scheme, and only a caller who knows there are two asks for the other. A default that is safe was preferred to a safe option, as in ADR 0012.

2. Keep both schemes for writing, and make Agile the default.
   The obvious call gives the stronger protection. The weaker scheme stays one keyword away, and its writer stays in the library to be maintained and tested. Its only use is a file for a reader that cannot open Agile encryption, which for Excel means a version from before 2010. That was not judged worth keeping a way to write AES-128 in ECB mode without integrity.

3. Write Agile only, and read both.
   Chosen. xlsxrb writes one scheme, the stronger one, and there is no option that weakens a file. Files that already exist with Standard encryption stay readable. A caller who reads such a file and writes it again gets an Agile file.

4. Write and read Agile only.
   The least code. Files encrypted by Excel 2007, and by tools that write Standard, could no longer be opened. Reading a weakly encrypted file does not weaken anything: the harm of a weak scheme is in the files written with it. Refusing to read would cost callers their existing files and protect nobody.

5. Accept `encryption_mode: :standard` and write Agile.
   No call breaks. A caller who named the weaker scheme did so for a reader that needs it, and would get a file that reader cannot open, with nothing at the call to say so. Under Principle 3 an option that is no longer supported is an error.

Two smaller questions were settled with these.

* A caller who passes `password:` without `encryption_mode:` is not warned before the default changes. The call keeps working, the password is the same, and every Excel since 2010 opens the result, so there is nothing for the caller to change. A warning on every encrypted write would ask most callers to act on something that does not affect them. The change is announced in the `Deprecated` section of `CHANGELOG.md`, as the deprecation policy provides for a change of result that the call cannot detect.
* How reading Standard is tested once the library no longer writes it. A fixed file from another implementation shows that xlsxrb reads what others write, which a file xlsxrb wrote itself cannot show. Code that encrypts with Standard is also needed, to build the malformed and boundary cases a fixed file does not contain. Both are kept: fixed files under `test/fixtures/encryption/`, and the Standard encryption code moved from `lib/` to the test support code when writing is removed.

---

## Guiding Principles

1. What xlsxrb writes is protected as well as the format allows:
   A file written with a password uses the strongest scheme that the readers in use can open. There is no option that makes it weaker.
2. Reading is wider than writing:
   xlsxrb reads every scheme it can decrypt. A weak scheme is a reason not to write a file, not a reason to refuse one that exists.
3. An option that is not supported is an error:
   A scheme the caller names and xlsxrb does not write is rejected at the call. It is not replaced by another.
4. Protection a file claims is checked:
   Where the scheme carries a check that the file is unchanged, the reader makes the check and refuses the file if it fails.
5. Every scheme has a file from another implementation:
   Each scheme xlsxrb reads or writes is tested against a file that xlsx software other than xlsxrb wrote, so that agreement with itself is not taken for agreement with the format.

---

## Decision

* From 0.3.0, `Xlsxrb.write` with `password:` writes Agile encryption: AES-256 in CBC mode, SHA-512 with 100,000 iterations, and an HMAC-SHA-512 of the encrypted package. In the 0.2 series the default stays Standard, and `encryption_mode: :agile` selects Agile.
* `encryption_mode: :standard` with a password warns in the 0.2 series and raises `ArgumentError` from 0.3.0. `encryption_mode: :agile` is accepted in both and writes the same file as the default does from 0.3.0.
* `Xlsxrb.read` and `Xlsxrb.load` read files with Standard and with Agile encryption, in every release.
* A file with Agile encryption whose HMAC does not match the encrypted package, or that has no HMAC, is refused with `Xlsxrb::DecryptionError`. Files written by xlsxrb 0.2.1 and earlier are checked against the HMAC those versions wrote.
* A file read with Standard encryption and written again with a password is written with Agile encryption.

This ADR does not change ADR 0007: the entry points and their `password:` argument stay as they are.

---

## Consequences

### Positive

* From 0.3.0 the obvious call gives the stronger scheme, and no call writes the weaker one.
* A file written with a password can be checked for changes by any reader that verifies the HMAC, and xlsxrb is such a reader.
* One encryption writer is maintained instead of two.
* Existing files stay readable whichever scheme they use.

### Negative / Trade-offs

* A file written by xlsxrb 0.3.0 with a password cannot be opened by software that reads only Standard encryption, such as Excel 2007 as first released. A caller who needs such a file has to encrypt it with another tool.
* If a problem is found in the Agile writer, there is no other scheme to switch to until it is fixed.
* Callers that pass `encryption_mode: :standard` have to change the call before 0.3.0.
* Callers that rely on the default get a file with a different scheme in 0.3.0 without having been warned at the call.
* xlsxrb refuses an Agile file whose HMAC does not match, while Excel opens it. A file from a tool that writes a wrong HMAC opens in Excel and not in xlsxrb. xlsxrb's own earlier files are the one case provided for.
* Reading Standard is tested with fixed files and with encryption code that is no longer part of the library, so that code has to be kept working for the tests alone.
* Excel was checked on one version and one platform: it opened files through automation, and one workbook it saved with a password was examined. The test files written by Excel are two workbooks from the tests of msoffcrypto-tool, both from Excel 2016 for Mac and both with Agile encryption; for Standard encryption the tests rest on a file from Apache POI.
