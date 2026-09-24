using System.Diagnostics;
using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Cli
{
    internal static class CliCommands
    {
        internal static int Sheets(string path, TextWriter stdout, TextWriter stderr, string? password = null)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet: null, password);
                for (int i = 0; i < reader.SheetCount; i++)
                {
                    stdout.Write(i.ToString(CultureInfo.InvariantCulture));
                    stdout.Write('\t');
                    stdout.WriteLine(reader.SheetNameAt(i));
                }
                return 0;
            }, stderr);
        }

        private static readonly string[] _validFormats = ["xlsx", "xlsb", "xls", "csv"];

        internal static int Convert(string path, string? sheet, string? output, string? format, char delimiter, Stream stdout, TextWriter stderr, Action<int>? onProgress = null, string? password = null, char? inputDelimiter = null)
        {
            return Execute(() =>
            {
                string resolvedFormat = ResolveFormat(format, output);
                ThrowIfOutputIsADirectory(output);
                using IExcelRowReader reader = Open(path, sheet, password, inputDelimiter);

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

        private static void ThrowIfOutputIsADirectory(string? output)
        {
            if (output is not null && Directory.Exists(output))
            {
                throw new ArgumentException($"--output '{output}' is a directory, not a file.", nameof(output));
            }
        }

        private static void WriteCsv(IExcelRowReader reader, Stream target, bool leaveOpen, char delimiter, Action<int>? onProgress)
        {
            using CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(target, leaveOpen, new CsvWriterOptions { Delimiter = AsciiByte(delimiter, "delimiter") });
            WriteRows<CsvWorkbookWriter, CsvSheetWriter, CsvRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXlsx(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(target, leaveOpen);
            WriteRows<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXlsb(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(target, leaveOpen, new XlsbWriterOptions { Date1904 = reader.IsDate1904 });
            WriteRows<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(workbook, reader, onProgress);
        }

        private static void WriteXls(IExcelRowReader reader, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(target, leaveOpen, date1904: reader.IsDate1904);
            WriteRows<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(workbook, reader, onProgress);
        }

        private const int ProgressInterval = 500;

        private static void WriteRows<TWorkbook, TSheet, TRow>(TWorkbook workbook, IExcelRowReader reader, Action<int>? onProgress)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TSheet sheetWriter = workbook.AddSheet(reader.SheetName);

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
                            row.Write(!cellValue.Value.IsEmpty && cellValue.Value[0] != (byte)'0');
                            break;
                        case CellType.Number:
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

        internal static int Schema(string path, string? sheet, int headerRow, int sampleSize, TextWriter stdout, TextWriter stderr, string? password = null, char? inputDelimiter = null)
        {
            return Execute(() =>
            {
                using IExcelRowReader reader = Open(path, sheet, password, inputDelimiter);

                foreach (ExcelColumnSchema column in Excel.InferSchema(reader, headerRow, sampleSize))
                {
                    stdout.Write(column.Index.ToString(CultureInfo.InvariantCulture));
                    stdout.Write('\t');
                    stdout.Write(column.Name ?? string.Empty);
                    stdout.Write('\t');
                    stdout.Write(column.Type.ToString());
                    stdout.WriteLine(column.IsNullable ? "?" : string.Empty);
                }
                return 0;
            }, stderr);
        }

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
                                           or InvalidOperationException)
            {
                stderr.WriteLine(exception.Message);
                return 1;
            }
        }

        internal static byte AsciiByte(char value, string parameterName)
        {
            if (value > (char)127)
            {
                throw new ArgumentException($"--{parameterName} must be an ASCII character; got '{value}'.", parameterName);
            }
            return (byte)value;
        }

        internal static IExcelRowReader Open(string path, string? sheet, string? password = null, char? inputDelimiter = null)
        {
            bool isCsv = string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);
            CsvReaderOptions csv = inputDelimiter is char delimiter
                ? CsvReaderOptions.Default with { Delimiter = AsciiByte(delimiter, "input-delimiter") }
                : CsvReaderOptions.Default with { SniffDialect = true };
            ExcelReaderOptions options = new() { Csv = csv };
            if (password is not null)
            {
                options = options with { Password = password };
            }
            IExcelRowReader reader = Excel.Open(path, isCsv ? ExcelFileFormat.Csv : ExcelFileFormat.Unknown, options);

            if (sheet is null)
            {
                return reader;
            }

            try
            {
                if (reader.TryMoveToSheet(sheet))
                {
                    return reader;
                }
                if (int.TryParse(sheet, CultureInfo.InvariantCulture, out int index))
                {
                    reader.MoveToSheet(index);
                    return reader;
                }
                throw new ArgumentException($"no sheet named '{sheet}' in {path}.", nameof(sheet));
            }
            catch
            {
                reader.Dispose();
                throw;
            }
        }
    }
}
