using System.Globalization;
using System.Text;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Parser.Internal;
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

        internal static void XlsxDifferential(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() => CompareReaders(
                bytes,
                static b => Excel.FromXlsx(new MemoryStream(b, writable: false), leaveOpen: false, Limits),
                static b => Excel.FromXlsx(new ReadOnlyMemory<byte>(b), Limits),
                "stream",
                "memory"));
        }

        internal static void XlsbDifferential(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() => CompareReaders(
                bytes,
                static b => Excel.FromXlsb(new MemoryStream(b, writable: false), leaveOpen: false, Limits),
                static b => Excel.FromXlsb(new ReadOnlyMemory<byte>(b), Limits),
                "stream",
                "memory"));
        }

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

        internal static void Encrypted(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            FuzzOracle.Guard(() =>
            {
                using IExcelRowReader reader = Excel.Open(bytes, EncryptedLimits);
                DrainAllSheets(reader);
            });
        }

        internal static int OpenEncryptedSeedForSelfCheck(ReadOnlySpan<byte> data)
        {
            byte[] bytes = data.ToArray();
            using IExcelRowReader reader = Excel.Open(bytes, EncryptedLimits);
            int rows = 0;
            int sheets = reader.SheetCount;
            for (int i = 0; i < sheets; i++)
            {
                reader.MoveToSheet(i);
                using IExcelRowEnumerator e = reader.GetEnumerator();
                while (e.MoveNext())
                {
                    rows++;
                }
            }
            return rows;
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

        private sealed class FuzzRow
        {
            public string? Name { get; set; }
            public int Age { get; set; }
            public string? Note { get; set; }
        }

        internal static void CsvParallel(ReadOnlySpan<byte> data)
        {
            if (data.Length < 2)
            {
                return;
            }
            byte[] header = "Name,Age,Note\n"u8.ToArray();
            byte[] bytes = new byte[header.Length + data.Length];
            header.CopyTo(bytes, 0);
            data.CopyTo(bytes.AsSpan(header.Length));
            int chunkSize = 1 + (data[0] % 64);

            FuzzOracle.Guard(() =>
            {
                List<string>? sequential = null;
                Exception? sequentialFailure = null;
                try
                {
                    sequential = ParseSequential(bytes);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    sequentialFailure = ex;
                }

                List<string>? parallel = null;
                Exception? parallelFailure = null;
                try
                {
                    parallel = ParseParallel(bytes, chunkSize);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    parallelFailure = ex;
                }

                if (sequentialFailure is not null || parallelFailure is not null)
                {
                    if (sequentialFailure?.GetType() != parallelFailure?.GetType())
                    {
                        throw new InvalidOperationException(
                            $"Oracle divergence: sequential threw {sequentialFailure?.GetType().Name ?? "nothing"}, " +
                            $"parallel threw {parallelFailure?.GetType().Name ?? "nothing"}.");
                    }
                    return;
                }

                if (!sequential!.SequenceEqual(parallel!, StringComparer.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Oracle divergence: sequential yielded {sequential!.Count} rows, parallel yielded {parallel!.Count}.");
                }
            });
        }

        private static List<string> ParseSequential(byte[] bytes)
        {
            using CsvReader reader = Excel.FromCsv(bytes, CsvLimits);
            var rows = new List<string>();
            foreach (FuzzRow row in ExcelParser.FromAttributes<FuzzRow>().Parse(reader))
            {
                rows.Add(Render(row));
            }
            return rows;
        }

        private static List<string> ParseParallel(byte[] bytes, int chunkSize)
        {
            var rows = new List<string>();
            IAsyncEnumerable<FuzzRow> source = ParallelCsvFactory.CreateWithChunkSize<FuzzRow>(
                bytes.AsMemory(), degreeOfParallelism: 4, chunkSize, CsvLimits, config: null, CancellationToken.None);
            IAsyncEnumerator<FuzzRow> e = source.GetAsyncEnumerator(CancellationToken.None);
            try
            {
                while (true)
                {
                    ValueTask<bool> moveNext = e.MoveNextAsync();
                    bool hasNext = moveNext.AsTask().GetAwaiter().GetResult();
                    if (!hasNext)
                    {
                        break;
                    }
                    rows.Add(Render(e.Current));
                }
            }
            finally
            {
                e.DisposeAsync().AsTask().GetAwaiter().GetResult();
            }
            return rows;
        }

        private static string Render(FuzzRow row)
        {
            return string.Create(CultureInfo.InvariantCulture, $"{row.Name}{row.Age}{row.Note}");
        }

        internal static int OpenDifferentialSeedForSelfCheck(ReadOnlySpan<byte> data, bool xlsb)
        {
            byte[] bytes = data.ToArray();
            List<string> stream = RenderAllSheets(xlsb
                ? Excel.FromXlsb(new MemoryStream(bytes, writable: false), leaveOpen: false, Limits)
                : Excel.FromXlsx(new MemoryStream(bytes, writable: false), leaveOpen: false, Limits));
            List<string> memory = RenderAllSheets(xlsb
                ? Excel.FromXlsb(new ReadOnlyMemory<byte>(bytes), Limits)
                : Excel.FromXlsx(new ReadOnlyMemory<byte>(bytes), Limits));
            AssertSameRows(stream, memory, "stream", "memory");
            return stream.Count;
        }

        internal static void AssertSameRowsSelfCheck()
        {
            try
            {
                AssertSameRows(["0:1=a@-1;"], ["0:1=b@-1;"], "left", "right");
            }
            catch (InvalidOperationException)
            {
                AssertSameRows(["0:1=a@-1;"], ["0:1=a@-1;"], "left", "right");
                return;
            }
            throw new InvalidOperationException("AssertSameRows accepted two different rows.");
        }

        private static void CompareReaders(
            byte[] bytes,
            Func<byte[], IExcelRowReader> left,
            Func<byte[], IExcelRowReader> right,
            string leftName,
            string rightName)
        {
            List<string>? leftRows = null;
            Exception? leftFailure = null;
            try
            {
                leftRows = RenderAllSheets(left(bytes));
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                leftFailure = ex;
            }

            List<string>? rightRows = null;
            Exception? rightFailure = null;
            try
            {
                rightRows = RenderAllSheets(right(bytes));
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                rightFailure = ex;
            }

            if (leftFailure is not null || rightFailure is not null)
            {
                if (leftFailure?.GetType() != rightFailure?.GetType())
                {
                    throw new InvalidOperationException(
                        $"Oracle divergence: {leftName} threw {leftFailure?.GetType().Name ?? "nothing"}, " +
                        $"{rightName} threw {rightFailure?.GetType().Name ?? "nothing"}.");
                }
                return;
            }

            AssertSameRows(leftRows!, rightRows!, leftName, rightName);
        }

        private static void AssertSameRows(List<string> left, List<string> right, string leftName, string rightName)
        {
            int common = Math.Min(left.Count, right.Count);
            for (int i = 0; i < common; i++)
            {
                if (!string.Equals(left[i], right[i], StringComparison.Ordinal))
                {
                    throw new InvalidOperationException(
                        $"Oracle divergence at row {i}: {leftName} read \"{left[i]}\", {rightName} read \"{right[i]}\".");
                }
            }

            if (left.Count != right.Count)
            {
                throw new InvalidOperationException(
                    $"Oracle divergence: {leftName} yielded {left.Count} rows, {rightName} yielded {right.Count}.");
            }
        }

        private static List<string> RenderAllSheets(IExcelRowReader reader)
        {
            using (reader)
            {
                var rows = new List<string>();
                int sheets = reader.SheetCount;
                for (int i = 0; i < sheets; i++)
                {
                    reader.MoveToSheet(i);
                    rows.Add(string.Create(CultureInfo.InvariantCulture, $"#sheet{i}"));
                    using IExcelRowEnumerator e = reader.GetEnumerator();
                    while (e.MoveNext())
                    {
                        rows.Add(RenderRow(e.Current, reader.IsDate1904));
                    }
                }
                return rows;
            }
        }

        private static string RenderRow(Row row, bool isDate1904)
        {
            var sb = new StringBuilder();
            foreach (RowCell rowCell in row.Cells)
            {
                Cell cell = rowCell.Value;
                long ticks = cell.TryGetDateTime(isDate1904, out DateTime date) ? date.Ticks : -1;
                sb.Append(CultureInfo.InvariantCulture, $"{rowCell.ColumnIndex}:{(int)cell.Type}={cell.GetString()}@{ticks};");
            }
            return sb.ToString();
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

            int probe = Math.Min(columns, 512);
            for (int c = 0; c < probe; c++)
            {
                _ = row[c].Type;
            }
        }
    }
}
