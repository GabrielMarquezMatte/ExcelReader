using System.Diagnostics;
using System.Globalization;
using System.Linq;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using ExcelReader.Core.Writer;

namespace ExcelReader.Cli
{
    /// <summary>
    /// The command bodies, as plain functions over explicit writers — this is the layer the tests
    /// drive. Nothing here touches <c>Console</c> or any other process-global state, so the tests
    /// run in parallel without interfering with each other.
    /// </summary>
    internal static class CliCommands
    {
        internal static int Sheets(string path, TextWriter stdout, TextWriter stderr, string? password = null)
        {
            return Sheets(path, (index, name) =>
            {
                stdout.Write(index.ToString(CultureInfo.InvariantCulture));
                stdout.Write('\t');
                stdout.WriteLine(name);
            }, stderr, password);
        }

        /// <summary>
        /// Same command, reported through a callback instead of a fixed tab-separated line - the shape
        /// Commands.cs's interactive (table-rendering) path needs, without a second implementation of
        /// the sheet-listing loop or its error handling.
        /// </summary>
        internal static int Sheets(string path, Action<int, string> onSheet, TextWriter stderr, string? password = null)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet: null, password);
                for (int i = 0; i < reader.SheetCount; i++)
                {
                    onSheet(i, reader.SheetNameAt(i));
                }
                return 0;
            }, stderr);
        }

        // The extensions/format names this command understands, and the only values --format
        // accepts. Order here is the order every "expected one of ..." error message lists them in.
        // Not compile-time tied to anything else that names these four: Commands.cs's --format/path
        // XML doc comments spell them out as prose for --help, and the switch in Convert below has
        // its own case per format (guarded by an UnreachableException default, so THAT one fails
        // loudly if it drifts). Adding a 5th format means updating all three by hand.
        private static readonly string[] _validFormats = ["xlsx", "xlsb", "xls", "csv"];

        internal static int Convert(string path, string? sheet, string? output, string? format, char delimiter, Stream stdout, TextWriter stderr, Action<int>? onProgress = null, string? password = null)
        {
            return Execute(() =>
            {
                string resolvedFormat = ResolveFormat(format, output);
                ThrowIfOutputIsADirectory(output);
                using IExcelRowReader reader = Open(path, sheet, password);

                bool leaveOpen = output is null;
                Stream target = leaveOpen
                    ? stdout
                    : new FileStream(output!, FileMode.Create, FileAccess.Write, FileShare.None);
                try
                {
                    switch (resolvedFormat)
                    {
                        case "csv":
                            WriteCsv(reader, target, leaveOpen, delimiter, onProgress);
                            break;
                        case "xlsx":
                            WriteXlsx(reader, target, leaveOpen, onProgress);
                            break;
                        case "xlsb":
                            WriteXlsb(reader, target, leaveOpen, onProgress);
                            break;
                        case "xls":
                            WriteXls(reader, target, leaveOpen, onProgress);
                            break;
                        default:
                            // Unreachable: ResolveFormat only ever returns one of the cases above.
                            throw new UnreachableException($"unresolved format '{resolvedFormat}'.");
                    }
                }
                finally
                {
                    if (!leaveOpen)
                    {
                        target.Dispose();
                    }
                }
                return 0;
            }, stderr);
        }

        /// <summary>
        /// Picks the workbook format to write: <paramref name="format"/> wins when given; otherwise
        /// it's inferred from <paramref name="output"/>'s extension; with neither (writing to
        /// stdout with no override), it's CSV - the one format every shell can already consume.
        /// </summary>
        /// <exception cref="ArgumentException"><paramref name="format"/> isn't one of
        /// <see cref="_validFormats"/>, or <paramref name="output"/>'s extension isn't either.</exception>
        internal static string ResolveFormat(string? format, string? output)
        {
            if (format is not null)
            {
                string normalized = format.ToLowerInvariant();
                if (Array.IndexOf(_validFormats, normalized) < 0)
                {
                    throw new ArgumentException(
                        $"unknown format '{format}'; expected one of {string.Join(", ", _validFormats)}.", nameof(format));
                }
                return normalized;
            }

            if (output is not null)
            {
                string extension = Path.GetExtension(output).TrimStart('.').ToLowerInvariant();
                if (Array.IndexOf(_validFormats, extension) < 0)
                {
                    // Path.GetExtension("noext") and Path.GetExtension("some/dir/") both return "" -
                    // reporting that as "unrecognized extension '.'" reads as if the user typed a
                    // bare dot, when in fact they gave none. Name the actual problem instead.
                    string problem = extension.Length == 0
                        ? $"--output '{output}' has no file extension"
                        : $"unrecognized output extension '.{extension}'";
                    throw new ArgumentException(
                        $"{problem}; expected one of {string.Join(", ", _validFormats.Select(static f => "." + f))}, or pass --format explicitly.",
                        nameof(output));
                }
                return extension;
            }

            return "csv";
        }

        /// <summary>
        /// A directory whose name happens to end in a recognized extension (e.g. a folder literally
        /// named "backup.xlsx") resolves a format fine, then fails <see cref="FileStream"/> with
        /// "Access to the path is denied" - true, but it reads as a permissions problem rather than
        /// naming the actual mistake.
        /// </summary>
        /// <exception cref="ArgumentException"><paramref name="output"/> names an existing directory.</exception>
        private static void ThrowIfOutputIsADirectory(string? output)
        {
            if (output is not null && Directory.Exists(output))
            {
                throw new ArgumentException($"--output '{output}' is a directory, not a file.", nameof(output));
            }
        }

        private static void WriteCsv(IExcelRowReader reader, Stream target, bool leaveOpen, char delimiter, Action<int>? onProgress)
        {
            using CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(target, leaveOpen, new CsvWriterOptions { Delimiter = (byte)delimiter });
            WriteRows<CsvWorkbookWriter, CsvSheetWriter, CsvRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXlsx(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(target, leaveOpen);
            WriteRows<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXlsb(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(target, leaveOpen, date1904: reader.IsDate1904);
            WriteRows<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXls(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(target, leaveOpen, date1904: reader.IsDate1904);
            WriteRows<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(workbook, reader, onProgress);
        }

        /// <summary>
        /// Copies every sampled row's cells across as text, one <see cref="IRowWriter.Write(string?)"/>
        /// call per column. Generic over the four workbook/sheet/row writer triples so the loop - the
        /// only part that actually varies by target - lives once; each format still gets its own
        /// <c>Create</c> call above, since their constructor parameters (date1904, compression, ...)
        /// differ.
        /// </summary>
        // How many rows between onProgress callbacks - frequent enough to look alive on a terminal
        // (~a handful of updates per second on typical row-write throughput), rare enough that the
        // callback (usually a Spectre.Console status refresh) never dominates the actual write.
        private const int ProgressInterval = 500;

        private static void WriteRows<TWorkbook, TSheet, TRow>(TWorkbook workbook, IExcelRowReader reader, Action<int>? onProgress)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            workbook.Start();
            using TSheet sheetWriter = workbook.AddSheet(reader.SheetName);
            sheetWriter.Start();

            int rowCount = 0;
            using IExcelRowEnumerator rows = reader.GetEnumerator();
            while (rows.MoveNext())
            {
                using TRow row = sheetWriter.StartRow();
                Row current = rows.Current;
                foreach (var cell in current.Cells)
                {
                    var cellValue = cell.Value;
                    switch (cellValue.Type)
                    {
                        case CellType.Boolean:
                            // Every format's Boolean cell stores a single raw byte, '0' or '1'
                            // (CellAccumulator.AddBool) - not the text "True"/"False" bool.TryParse
                            // expects, so this reads the byte directly rather than going through
                            // GetString()/Cell.TryParse<bool>. Writing a real bool cell (not text)
                            // preserves the source's type in the converted workbook.
                            row.Write(!cellValue.Value.IsEmpty && cellValue.Value[0] != (byte)'0');
                            break;
                        case CellType.Number:
                            // Writing a real numeric cell (not text) preserves the source's type -
                            // a text fallback would make every number a left-aligned string cell in
                            // the converted workbook.
                            if (cellValue.TryGetDouble(out double number))
                            {
                                row.Write(number);
                            }
                            else
                            {
                                row.Write(cellValue.GetString());
                            }
                            break;
                        case CellType.Date:
                            var dateValue = cellValue.TryGetDateTime(out var date) ? date : throw new InvalidOperationException($"cell {cell.ColumnIndex} is a date but TryGetDateTime failed to parse it.");
                            row.Write(dateValue);
                            break;
                        default:
                            row.Write(cellValue.GetString());
                            break;
                    }
                }
                rowCount++;
                if (onProgress is not null && rowCount % ProgressInterval == 0)
                {
                    onProgress(rowCount);
                }
            }
            onProgress?.Invoke(rowCount);
            sheetWriter.End();
            workbook.End();
        }

        internal static int Schema(string path, string? sheet, int headerRow, int sampleSize, TextWriter stdout, TextWriter stderr, string? password = null)
        {
            return Schema(path, sheet, headerRow, sampleSize, column =>
            {
                stdout.Write(column.Index.ToString(CultureInfo.InvariantCulture));
                stdout.Write('\t');
                // A null name means the column is addressable only by index - an empty field,
                // never the literal "null", so the output stays machine-parseable.
                stdout.Write(column.Name ?? string.Empty);
                stdout.Write('\t');
                stdout.Write(column.Type.ToString());
                stdout.WriteLine(column.IsNullable ? "?" : string.Empty);
            }, stderr, password);
        }

        /// <summary>
        /// Same command, reported through a callback instead of a fixed tab-separated line - see the
        /// <see cref="Sheets(string, Action{int, string}, TextWriter, string?)"/> overload for why.
        /// </summary>
        internal static int Schema(string path, string? sheet, int headerRow, int sampleSize, Action<ExcelColumnSchema> onColumn, TextWriter stderr, string? password = null)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet, password);

                foreach (ExcelColumnSchema column in Excel.InferSchema(reader, headerRow, sampleSize))
                {
                    onColumn(column);
                }
                return 0;
            }, stderr);
        }

        /// <summary>
        /// Runs <paramref name="body"/>, turning the failures a user can act on into exit code 1 plus
        /// a single stderr line. Anything not listed here is a bug and is deliberately left to
        /// propagate with its stack trace intact.
        /// </summary>
        internal static int Execute(Func<int> body, TextWriter stderr)
        {
            try
            {
                return body();
            }
            catch (Exception exception) when (exception is IOException
                                           or UnauthorizedAccessException
                                           or InvalidDataException
                                           or ArgumentException
                                           or NotSupportedException
                                           or ExcelLimitExceededException
                                           // A legacy XLS target rejects sheets wider than 256
                                           // columns (BIFF8's own limit) - a real user-facing
                                           // "this workbook doesn't fit that format" error, not a bug.
                                           or InvalidOperationException)
            {
                stderr.WriteLine(exception.Message);
                return 1;
            }
        }

        /// <summary>
        /// Opens <paramref name="path"/> and selects <paramref name="sheet"/>, which is either a
        /// zero-based index or a sheet name. CSV is opened through its own factory because
        /// <see cref="Excel.Open(string, ExcelReaderOptions?)"/> deliberately does not sniff it, and
        /// <paramref name="password"/> is ignored for it since CSV is never encrypted.
        /// </summary>
        internal static IExcelRowReader Open(string path, string? sheet, string? password = null)
        {
            bool isCsv = string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);
            ExcelReaderOptions? options = password is null ? null : new ExcelReaderOptions { Password = password };
            IExcelRowReader reader = isCsv
                ? Excel.FromCsvFile(path)
                : Excel.Open(path, options);

            if (sheet is null)
            {
                return reader;
            }

            try
            {
                if (int.TryParse(sheet, CultureInfo.InvariantCulture, out int index))
                {
                    reader.MoveToSheet(index);
                    return reader;
                }
                if (!reader.TryMoveToSheet(sheet))
                {
                    throw new ArgumentException($"no sheet named '{sheet}' in {path}.", nameof(sheet));
                }
                return reader;
            }
            catch
            {
                reader.Dispose();
                throw;
            }
        }
    }
}
