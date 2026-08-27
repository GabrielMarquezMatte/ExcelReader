using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;

namespace ExcelReader.Fuzz
{
    /// <summary>
    /// One entry point per fuzz target. Each takes arbitrary bytes, drives a reader over them, and
    /// lets <see cref="FuzzOracle"/> decide whether any resulting exception is acceptable.
    /// </summary>
    internal static class Harnesses
    {
        // Deliberately tight limits: they keep a single input fast (the engine runs millions of them)
        // and, more importantly, they put the limit checks themselves on the hot path so the fuzzer
        // actually explores the guard code rather than only the happy path.
        private static readonly ExcelReaderOptions Limits = new()
        {
            MaxCellBytes = 1 << 20,
            MaxSharedStringBytes = 1 << 22,
        };

        private static readonly CsvReaderOptions CsvLimits = new()
        {
            MaxCellBytes = 1 << 20,
        };

        private static readonly ExcelReaderOptions EncryptedLimits = new()
        {
            MaxCellBytes = 1 << 20,
            MaxSharedStringBytes = 1 << 22,
            Password = "hunter2",
            // Keep derivation cheap: the engine runs millions of inputs, and a real spin count would make
            // each one take ~100ms.
            MaxPasswordSpinCount = 1_000,
        };

        internal static void Xlsx(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using var ms = new MemoryStream(bytes, writable: false);
                using XlsxReader reader = Excel.FromXlsx(ms, leaveOpen: true, Limits);
                DrainAllSheets(reader);
            });
        }

        // The in-memory ZIP path (ZipMemoryIndex) is a different container parser from the
        // Stream/ZipArchive one above, so it gets its own target rather than sharing a corpus.
        internal static void XlsxMemory(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using XlsxReader reader = Excel.FromXlsx(new ReadOnlyMemory<byte>(bytes), Limits);
                DrainAllSheets(reader);
            });
        }

        internal static void Xlsb(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using var ms = new MemoryStream(bytes, writable: false);
                using XlsbReader reader = Excel.FromXlsb(ms, leaveOpen: true, Limits);
                DrainAllSheets(reader);
            });
        }

        internal static void XlsbMemory(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using XlsbReader reader = Excel.FromXlsb(new ReadOnlyMemory<byte>(bytes), Limits);
                DrainAllSheets(reader);
            });
        }

        // BIFF8 inside an OLE compound file — an entirely hand-rolled container parser, and the one
        // with the most pointer-like structures (sector chains, FAT/miniFAT) reachable from bytes.
        internal static void Xls(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using var ms = new MemoryStream(bytes, writable: false);
                using XlsReader reader = Excel.FromXls(ms, leaveOpen: true, Limits);
                DrainAllSheets(reader);
            });
        }

        // The encrypted container is a third container parser (CFB directory + EncryptionInfo descriptor)
        // layered under the ZIP one, and it runs BEFORE any password check - so it gets its own target.
        // The password is fixed and correct for the seed, so mutations explore the parsers rather than
        // dead-ending on a verifier mismatch.
        internal static void Encrypted(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using IExcelRowReader reader = Excel.Open(bytes, EncryptedLimits);
                DrainAllSheets(reader);
            });
        }

        internal static void Csv(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using var ms = new MemoryStream(bytes, writable: false);
                using CsvReader reader = Excel.FromCsv(ms, leaveOpen: true, CsvLimits);
                DrainRows(reader);
            });
        }

        // Dialect detection runs over untrusted bytes before any reader is constructed, so it is its
        // own attack surface — and unlike the readers it has no stream to bound it.
        internal static void CsvSniff(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                CsvDialect dialect = CsvSniffer.Detect(bytes);
                using var ms = new MemoryStream(bytes, writable: false);
                using CsvReader reader = Excel.FromCsv(ms, leaveOpen: true, CsvLimits.WithDialect(dialect));
                DrainRows(reader);
            });
        }

        private static void DrainAllSheets(IExcelRowReader reader)
        {
            int sheets = reader.SheetCount;
            for (int i = 0; i < sheets; i++)
            {
                reader.MoveToSheet(i);
                DrainRows(reader);
            }
        }

        private static void DrainRows(IExcelRowReader reader)
        {
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            while (rows.MoveNext())
            {
                TouchRow(rows.Current, reader.IsDate1904);
            }
        }

        private static void DrainRows(CsvReader reader)
        {
            using CsvReader.Enumerator rows = reader.GetEnumerator();
            while (rows.MoveNext())
            {
                TouchRow(rows.Current, isDate1904: false);
            }
        }

        // Materializing every cell is the point: Row/Cell hand out spans sliced from reader-owned
        // buffers using offsets decoded from the input, so a bad offset only becomes observable once
        // something actually reads through it.
        private static void TouchRow(Row row, bool isDate1904)
        {
            int columns = row.ColumnCount;
            Span<byte> scratch = stackalloc byte[64];
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                _ = cell.Type;
                _ = cell.StyleIndex;
                _ = cell.Value.Length;
                _ = cell.GetString();
                _ = cell.TryGetDouble(out _);
                _ = cell.TryGetDateTime(isDate1904, out _);
                _ = cell.TryParse<int>(CultureInfo.InvariantCulture, out _);
                _ = cell.TryFormat(scratch, out _);
            }

            // Indexer path: a separate binary search over the same descriptors as the enumerator
            // above, bounded so a corrupt ColumnCount cannot turn this into a near-infinite loop.
            int probe = Math.Min(columns, 512);
            for (int c = 0; c < probe; c++)
            {
                _ = row[c].Type;
            }
        }
    }
}
