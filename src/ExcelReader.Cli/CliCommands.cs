using System.Globalization;
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
        internal static int Sheets(string path, TextWriter stdout, TextWriter stderr)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet: null);
                for (int i = 0; i < reader.SheetCount; i++)
                {
                    stdout.Write(i.ToString(CultureInfo.InvariantCulture));
                    stdout.Write('\t');
                    stdout.WriteLine(reader.SheetNameAt(i));
                }
                return 0;
            }, stderr);
        }

        internal static int Convert(string path, string? sheet, string? output, char delimiter, Stream stdout, TextWriter stderr)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet);

                Stream target = output is null
                    ? stdout
                    : new FileStream(output, FileMode.Create, FileAccess.Write, FileShare.None);
                try
                {
                    CsvWriterOptions options = new() { Delimiter = (byte)delimiter };
                    using var workbook = CsvWorkbookWriter.Create(target, leaveOpen: output is null, options);

                    workbook.Start();
                    using var sheetWriter = workbook.AddSheet(reader.SheetName);
                    sheetWriter.Start();

                    using IExcelRowEnumerator rows = reader.GetEnumerator();
                    while (rows.MoveNext())
                    {
                        using var row = sheetWriter.StartRow();
                        Row current = rows.Current;
                        for (int column = 0; column < current.ColumnCount; column++)
                        {
                            // ponytail: GetString() allocates per non-shared cell. If profiling ever
                            // shows this dominates a large convert, the fix is one
                            // CsvRowWriter.Write(ReadOnlySpan<byte>) overload fed from Cell.Value -
                            // not anything in this file. Not doing it on spec.
                            row.Write(current[column].GetString());
                        }
                    }
                    sheetWriter.End();
                    workbook.End();
                }
                finally
                {
                    if (output is not null)
                    {
                        target.Dispose();
                    }
                }
                return 0;
            }, stderr);
        }

        internal static int Schema(string path, string? sheet, int headerRow, int sampleSize, TextWriter stdout, TextWriter stderr)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet);

                foreach (ExcelColumnSchema column in Excel.InferSchema(reader, headerRow, sampleSize))
                {
                    stdout.Write(column.Index.ToString(CultureInfo.InvariantCulture));
                    stdout.Write('\t');
                    // A null name means the column is addressable only by index - an empty field,
                    // never the literal "null", so the output stays machine-parseable.
                    stdout.Write(column.Name ?? string.Empty);
                    stdout.Write('\t');
                    stdout.Write(column.Type.ToString());
                    stdout.WriteLine(column.IsNullable ? "?" : string.Empty);
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
                                           or ExcelLimitExceededException)
            {
                stderr.WriteLine(exception.Message);
                return 1;
            }
        }

        /// <summary>
        /// Opens <paramref name="path"/> and selects <paramref name="sheet"/>, which is either a
        /// zero-based index or a sheet name. CSV is opened through its own factory because
        /// <see cref="Excel.Open(string, ExcelReaderOptions?)"/> deliberately does not sniff it.
        /// </summary>
        internal static IExcelRowReader Open(string path, string? sheet)
        {
            IExcelRowReader reader = string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase)
                ? Excel.FromCsvFile(path)
                : Excel.Open(path);

            if (sheet is null)
            {
                return reader;
            }

            try
            {
                if (int.TryParse(sheet, CultureInfo.InvariantCulture, out int index))
                {
                    reader.MoveToSheet(index);
                }
                else if (!reader.TryMoveToSheet(sheet))
                {
                    throw new ArgumentException($"no sheet named '{sheet}' in {path}.", nameof(sheet));
                }
            }
            catch
            {
                reader.Dispose();
                throw;
            }
            return reader;
        }
    }
}
