# Encrypted workbooks

Reading and writing password-protected packages.

## Encrypted workbooks

Password-protected `.xlsx`/`.xlsb`/`.xlsm` files open through the same entry points — the password
goes on the options, so every overload supports it:

```csharp
var options = new ExcelReaderOptions { Password = "hunter2" };
using IExcelRowReader reader = Excel.Open("protected.xlsx", options);
foreach (Row row in reader) { /* ... */ }
```

`ExcelEncryptionException.Reason` tells you what to do about a failure — only `PasswordRequired` and
`PasswordIncorrect` are worth re-prompting for:

```csharp
try
{
    using IExcelRowReader reader = Excel.Open(path, options);
}
catch (ExcelEncryptionException ex) when (ex.Reason is ExcelEncryptionReason.PasswordIncorrect)
{
    // Ask again. UnsupportedScheme and IntegrityFailure are terminal.
}
```

Supported: ECMA-376 agile encryption (Excel 2010+ — what Excel writes today when you set a
password) and ECMA-376 standard encryption (Excel 2007 — AES-ECB with SHA-1). **Not** supported:
RC4 CryptoAPI encryption, encrypted legacy `.xls`, and sheet/workbook *protection* passwords — a
different mechanism entirely, stored as hashes in the plaintext XML.

`Password` never appears in `ExcelReaderOptions.ToString()`. Note that a password supplied as a
`string` cannot be wiped from memory — .NET strings are immutable and movable — so the library zeroes
only the buffers it owns: the derivation buffer and the derived key.

**The `dataIntegrity` HMAC is not verified by default when reading from a stream.**
`ExcelReaderOptions.VerifyEncryptedIntegrity` defaults to `false`, because the HMAC covers the whole
encrypted package: checking it means a full pass over the file before the first row, which makes
time-to-first-row proportional to file size and defeats streaming. With it off, decryption still
fails loudly on corrupt ciphertext (the plaintext stops being a valid ZIP), but *targeted* tampering
by someone who can modify the file is not detected. Set it to `true` whenever the workbook comes from
somewhere you do not control and you care that it was not altered:

```csharp
var options = new ExcelReaderOptions
{
    Password = "hunter2",
    VerifyEncryptedIntegrity = true, // pay one full pass; reject tampered packages
};
```

The in-memory path (`Excel.Open(ReadOnlyMemory<byte>)` and friends) has already decrypted everything
by the time it returns, so it always verifies regardless of this setting. Standard encryption has no
HMAC field, so the setting does not apply to it.

Opening an encrypted workbook costs a fixed ~26 ms for key derivation, whatever the file size. After
that, reading is ~10% slower than reading the same workbook unencrypted — see
[Encrypted workbooks](../performance/benchmarks.md#encrypted-workbooks) in the benchmarks.

Writing an encrypted workbook is a second step: build the package with any writer, then wrap it with
`Excel.EncryptPackage`, which produces an agile-encrypted (ECMA-376 4.4) CFB container — AES-256-CBC,
SHA-512, 100,000 spin iterations, with a `dataIntegrity` HMAC.

```csharp
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsx;

using var plain = new MemoryStream();
await using (var workbook = XlsxWorkbookWriter.Create(plain, leaveOpen: true))
{
    // write sheets and rows as usual
}
plain.Position = 0;

using var destination = File.Create("secret.xlsx");
await Excel.EncryptPackageAsync(plain, destination, "hunter2");
```

The package stream must be seekable (it is read twice, so the container never has to be buffered in
memory); the destination does not. Path-to-path overloads exist for both the sync and async forms.
Encryption parameters are fixed at Excel's own defaults — there are no knobs — and only XLSX/XLSB
packages can be encrypted, matching what the reader can decrypt.

