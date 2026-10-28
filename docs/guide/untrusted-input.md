# Reading untrusted files

What one hostile workbook can cost you, which options bound it, and how to reject a file cleanly.
Read this before you accept spreadsheets from uploads, APIs or mail.

A spreadsheet is a container whose sizes, counts and offsets are written by whoever made the file.
The readers check every one of those values before it drives an allocation, and the options below
turn what is left — a file that is well-formed but simply enormous — into an
[`ExcelLimitExceededException`](xref:ExcelReader.Core.Reader.ExcelLimitExceededException) instead of an out-of-memory crash. The defaults suit
a desktop tool or a batch job reading files you trust. A server should set its own.

## The limits

All of them live on [`ExcelReaderOptions`](xref:ExcelReader.Core.Reader.ExcelReaderOptions):

| Option | Default | Bounds | Formats |
|---|---:|---|---|
| `MaxTotalDecompressedBytes` | 512 MiB | What one read of a sheet inflates or loads from its container, counting the workbook's metadata, styles, shared strings and decrypted package | XLSX, XLSB, XLS, encrypted |
| `MaxSharedStringBytes` | 128 MiB | The shared-string table, which is held in memory for the whole read | XLSX, XLSB, XLS |
| `MaxCellBytes` | 32 MiB | Each buffer that grows to hold one cell value, or one CSV record | all |
| `MaxZipEntries` | 65,536 | Entries in the ZIP directory | XLSX, XLSB |
| `MaxPasswordSpinCount` | 100,000 | Key-derivation iterations an agile-encrypted file may ask for | encrypted XLSX, XLSB |

Two limits are fixed by Excel's grid rather than configured: a cell past column XFD (16,384) throws
with `LimitName == "Columns"`, and no single buffer may exceed .NET's array size
(`LimitName == "ArrayMaxLength"`).

How they count:

- **`MaxTotalDecompressedBytes` is per read of a sheet.** The workbook's one-time loads (its
  metadata, styles and shared strings) count toward every read, and each read adds the sheet it
  inflates. Reading a sheet again starts from the one-time loads, so a workbook kept open and read
  repeatedly never runs out.
- **XLS has nothing to inflate, but it is not streamed:** the whole Workbook stream is read into
  memory, up to `MaxTotalDecompressedBytes`.
- **CSV has no total limit.** It streams, so the size of what you hand it is the bound; only
  `MaxCellBytes` applies.
- **Zero or a negative value turns a byte or entry limit off.** `MaxPasswordSpinCount` has no off
  switch. In the Python, Rust and C++ bindings and the C ABI, `0` means the default instead, and a
  limit cannot be turned off.
- **Reads running at once each have their own budget.** Ten uploads read together, or several sheets
  of one workbook read on different threads, can inflate ten times the budget between them.

## A profile for uploads

Size each limit for the largest file you legitimately expect, measured rather than guessed. For a
service taking spreadsheets of a few MB, something like:

<!-- snippet: untrusted-input-options -->
```csharp
var uploadOptions = new ExcelReaderOptions
{
    MaxTotalDecompressedBytes = 64L * 1024 * 1024, // everything one workbook may inflate
    MaxSharedStringBytes = 16L * 1024 * 1024,
    MaxCellBytes = 1024 * 1024,
    MaxZipEntries = 1_000,
};
```

Cap the upload itself as well, before it reaches the reader: the options bound what the file
expands to, not how many bytes you accept. [`Excel.Open`](xref:ExcelReader.Core.Reader.Excel.Open*) needs a seekable stream, so
buffer a request body to a file or to memory first.

Leave [`PrefetchDecompression`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.PrefetchDecompression) off: it adds a thread per read, which a busy server
cannot spare. For encrypted uploads, set `VerifyEncryptedIntegrity` so a tampered package is
rejected; see [Encrypted workbooks](encryption.md).

No option bounds wall-clock time. The work a read does grows with the bytes it inflates, which the
limits bound; for a deadline, use the async entry points and enumerators, which take a
`CancellationToken`.

## Rejecting a file

A limit can trip when the workbook opens, when you start reading a sheet (XLSX loads its shared
strings then), or halfway down it. Wrap the whole read, not just the open, and do not commit rows until the read
has finished:

<!-- snippet: untrusted-input-reject -->
```csharp
try
{
    using IExcelWorkbook workbook = Excel.Open(upload, options: uploadOptions);
    foreach (Row row in workbook.FirstSheet)
    {
        // stage the row; commit only once the loop has finished
    }
}
catch (ExcelLimitExceededException ex)
{
    Console.WriteLine($"Rejected: {ex.LimitName} reached {ex.Actual}, limit {ex.Limit}");
}
catch (Exception ex) when (IsRejectedFile(ex))
{
    Console.WriteLine($"Rejected: {ex.Message}");
}

static bool IsRejectedFile(Exception ex)
{
    return ex is InvalidDataException or EndOfStreamException or NotSupportedException
        or ExcelEncryptionException or ExcelParseException
    || ex.GetType() == typeof(ArgumentException)
    || ex.GetType() == typeof(InvalidOperationException);
}
```

`LimitName` is the option's name (or `Columns`, `ArrayMaxLength`); `Limit` and `Actual` are in that
limit's unit — bytes, entries or iterations.

That exception list is exactly what the fuzz suite accepts as a correct rejection of arbitrary
bytes. Anything else escaping a read of a file — an `IndexOutOfRangeException`, an
`OutOfMemoryException`, a hang — is a bug; please report it privately as described in
[SECURITY.md](../../SECURITY.md). See [Errors across languages](errors.md) for how these map onto the
bindings.

## Damaged packages

An XLSX or XLSB package is a ZIP archive, and each of its entries carries a CRC-32 of its data.
The reader checks it as the entry is read and throws `InvalidDataException` on a mismatch, so a file
cut short or altered in transit does not come back as plausible but wrong cells. A sheet is checked
when its last byte is read, which means the rows before the damage may already have been returned:
treat the exception as invalidating everything read from that file. The checksum guards against
accidents, not tampering, since whoever alters the data can recompute it.
[`ExcelReaderOptions.VerifyZipChecksums`](xref:ExcelReader.Core.Reader.ExcelReaderOptions.VerifyZipChecksums)
turns the check off.

## What the reader does not do

- **It evaluates nothing.** A formula cell reports the result Excel cached when it saved the file,
  so the value is whatever the file claims; recompute anything you rely on. Macros are never run
  and external links never followed.
- **It resolves no XML entities beyond the predefined ones.** Sheet XML goes through the library's
  own scanner, which decodes `&amp;`-style and numeric references and copies anything else through
  as text, so there is no entity expansion and no external fetch. The one `XmlReader`, for the
  encryption descriptor, prohibits DTDs.

## Writing untrusted text to CSV

The writers store text as text: in XLSX, XLSB and XLS a string that starts with `=` is still a
string. CSV has no cell types, and Excel opening a `.csv` treats a field starting with `=`, `+`,
`-` or `@` as a formula. [`CsvWorkbookWriter`](xref:ExcelReader.Core.Writer.Csv.CsvWorkbookWriter) writes your text unchanged, so if a person may open the
file in a spreadsheet program, neutralize those fields yourself:

<!-- snippet: untrusted-input-csv-injection -->
```csharp
static string Neutralize(string text)
{
    return text.Length > 0 && "=+-@\t\r".Contains(text[0]) ? "'" + text : text;
}

row.Write(Neutralize(comment));
```

The leading `'` changes the data, so do this only for files meant for people, not for files another
program parses back.
