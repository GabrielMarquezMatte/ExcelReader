using System.Diagnostics;
using System.Globalization;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Internal;
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
                using IExcelWorkbook workbook = Open(path, sheet: null, out _, password);
                for (int i = 0; i < workbook.SheetCount; i++)
                {
                    stdout.Write(i.ToString(CultureInfo.InvariantCulture));
                    stdout.Write('\t');
                    stdout.WriteLine(workbook.SheetAt(i).Name);
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
                using IExcelWorkbook workbook = Open(path, sheet, out IExcelSheet selected, password, inputDelimiter);

                if (output is null)
                {
                    WriteFormat(resolvedFormat, selected, stdout, leaveOpen: true, delimiter, onProgress);
                    return 0;
                }
                using StagedFile staged = new(output);
                WriteFormat(resolvedFormat, selected, staged.Stream, leaveOpen: true, delimiter, onProgress);
                staged.Commit();
                return 0;
            }, stderr);
        }

        private static void WriteFormat(string format, IExcelSheet selected, Stream target, bool leaveOpen, char delimiter, Action<int>? onProgress)
        {
            switch (format)
            {
                case "csv":
                    WriteCsv(selected, target, leaveOpen, delimiter, onProgress);
                    return;
                case "xlsx":
                    WriteXlsx(selected, target, leaveOpen, onProgress);
                    return;
                case "xlsb":
                    WriteXlsb(selected, target, leaveOpen, onProgress);
                    return;
                case "xls":
                    WriteXls(selected, target, leaveOpen, onProgress);
                    return;
                default:
                    throw new UnreachableException($"unresolved format '{format}'.");
            }
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

        private static void WriteCsv(IExcelSheet sheet, Stream target, bool leaveOpen, char delimiter, Action<int>? onProgress)
        {
            using CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(target, leaveOpen, new CsvWriterOptions { Delimiter = AsciiByte(delimiter, "delimiter") });
            WriteRows<CsvWorkbookWriter, CsvSheetWriter, CsvRowWriter>(workbook, sheet, onProgress);
        }

        private static void WriteXlsx(IExcelSheet sheet, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(target, leaveOpen, new XlsxWriterOptions { Date1904 = sheet.IsDate1904 });
            WriteRows<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(workbook, sheet, onProgress);
        }

        private static void WriteXlsb(IExcelSheet sheet, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(target, leaveOpen, new XlsbWriterOptions { Date1904 = sheet.IsDate1904 });
            WriteRows<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(workbook, sheet, onProgress);
        }

        private static void WriteXls(IExcelSheet sheet, Stream target, bool leaveOpen, Action<int>? onProgress)
        {
            using XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(target, leaveOpen, new XlsWriterOptions { Date1904 = sheet.IsDate1904 });
            WriteRows<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(workbook, sheet, onProgress);
        }

        private const int ProgressInterval = 500;

        private static void WriteRows<TWorkbook, TSheet, TRow>(TWorkbook workbook, IExcelSheet sheet, Action<int>? onProgress)
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            using TSheet sheetWriter = workbook.AddSheet(sheet.Name.Length == 0 ? "Sheet1" : sheet.Name);

            int rowCount = 0;
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            while (rows.MoveNext())
            {
                using TRow row = sheetWriter.StartRow();
                Row current = rows.Current;
                int nextColumn = 0;
                foreach (var cell in current.Cells)
                {
                    row.Skip(cell.ColumnIndex - nextColumn);
                    nextColumn = cell.ColumnIndex + 1;
                    var cellValue = cell.Value;
                    switch (cellValue.Type)
                    {
                        case CellType.Boolean:
                            var boolValue = ExcelCellReaders.Bool(in cellValue, default, out bool flag) ? flag : throw new InvalidOperationException($"cell {cell.ColumnIndex} is a boolean but holds '{cellValue.GetString()}'.");
                            row.Write(boolValue);
                            break;
                        case CellType.Number:
                            WriteNumber(row, in cellValue);
                            break;
                        case CellType.Date:
                            var dateValue = cellValue.TryGetDateTime(sheet.IsDate1904, out var date) ? date : throw new InvalidOperationException($"cell {cell.ColumnIndex} is a date but TryGetDateTime failed to parse it.");
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

        // Text-backed numbers keep their exact integer or decimal value; a double cannot tell apart
        // integers above 2^53.
        private static void WriteNumber<TRow>(TRow row, in Cell cell) where TRow : IRowWriter
        {
            if (!cell.Value.IsEmpty && cell.TryParse(CultureInfo.InvariantCulture, out long integer))
            {
                row.Write(integer);
                return;
            }
            bool hasDouble = cell.TryGetDouble(out double number);
            if (!cell.Value.IsEmpty && hasDouble && cell.TryParse(CultureInfo.InvariantCulture, out decimal exact)
                && double.Parse(exact.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture) == number)
            {
                row.Write(exact);
                return;
            }
            if (hasDouble)
            {
                row.Write(number);
                return;
            }
            row.Write(cell.GetString());
        }

        internal static int Schema(string path, string? sheet, int headerRow, int sampleSize, TextWriter stdout, TextWriter stderr, string? password = null, char? inputDelimiter = null)
        {
            return Execute(() =>
            {
                using IExcelWorkbook workbook = Open(path, sheet, out IExcelSheet selected, password, inputDelimiter);

                foreach (ExcelColumnSchema column in Excel.InferSchema(selected, new SchemaInferenceOptions { HeaderRow = headerRow, SampleSize = sampleSize }))
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

        internal static IExcelWorkbook Open(string path, string? sheet, out IExcelSheet selected, string? password = null, char? inputDelimiter = null)
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
            IExcelWorkbook workbook = Excel.Open(path, isCsv ? ExcelFileFormat.Csv : ExcelFileFormat.Unknown, options);
            try
            {
                selected = SelectSheet(workbook, sheet, path);
                return workbook;
            }
            catch
            {
                workbook.Dispose();
                throw;
            }
        }

        private static IExcelSheet SelectSheet(IExcelWorkbook workbook, string? sheet, string path)
        {
            if (sheet is null)
            {
                return workbook.SheetAt(0);
            }
            if (workbook.TryGetSheet(sheet, out IExcelSheet? named))
            {
                return named;
            }
            if (int.TryParse(sheet, CultureInfo.InvariantCulture, out int index))
            {
                return workbook.SheetAt(index);
            }
            throw new ArgumentException($"no sheet named '{sheet}' in {path}.", nameof(sheet));
        }
    }
}
