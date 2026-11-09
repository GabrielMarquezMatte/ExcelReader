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

        internal static int Preview(string path, string? sheet, int headerRow, int rows, int maxColumns, int maxWidth, bool fast, string? columns, TextWriter stdout, TextWriter stderr, string? password = null, char? inputDelimiter = null)
        {
            return Execute(() =>
            {
                RequirePositive(rows, "rows");
                RequirePositive(maxColumns, "max-columns");
                RequirePositive(maxWidth, "max-width");
                using IExcelWorkbook workbook = Open(path, sheet, out IExcelSheet selected, password, inputDelimiter);
                IReadOnlyList<ExcelColumnSchema> schema = Excel.InferSchema(selected, new SchemaInferenceOptions { HeaderRow = headerRow, ParseText = IsCsv(path) });
                List<int>? selection = ResolveColumns(columns, schema);
                HashSet<int>? wanted = selection is null ? null : [.. selection];
                int headColumns = (maxColumns + 1) / 2;
                int tailColumns = maxColumns - headColumns;

                int headCount = fast ? rows : (rows + 1) / 2;
                int tailCount = rows - headCount;
                List<string?[]> head = [];
                string?[][] tail = new string?[tailCount][];
                long total = 0;
                int columnCount = schema.Count;
                using IExcelRowEnumerator enumerator = selected.GetEnumerator();
                for (int i = 0; i < headerRow; i++)
                {
                    enumerator.MoveNext();
                }
                while (enumerator.MoveNext())
                {
                    Row row = enumerator.Current;
                    columnCount = Math.Max(columnCount, row.ColumnCount);
                    ColumnFilter filter = new(wanted, headColumns, columnCount - tailColumns);
                    if (total < headCount)
                    {
                        head.Add(Materialize(row, selected.IsDate1904, schema, filter));
                    }
                    if (total >= headCount && tailCount > 0)
                    {
                        tail[(total - headCount) % tailCount] = Materialize(row, selected.IsDate1904, schema, filter);
                    }
                    total++;
                    if (fast && total > rows)
                    {
                        break;
                    }
                }

                List<string?[]?> body = [.. head];
                long tailRegion = total - headCount;
                long kept = Math.Min(Math.Max(tailRegion, 0), tailCount);
                if (tailRegion > tailCount)
                {
                    body.Add(null);
                }
                for (long j = tailRegion - kept; j < tailRegion; j++)
                {
                    body.Add(tail[j % tailCount]);
                }
                string shownTotal = fast && total > rows ? "?" : Underscored(total);
                WriteTable(stdout, schema, shownTotal, selection ?? [.. Enumerable.Range(0, columnCount)], body, maxColumns, maxWidth);
                return 0;
            }, stderr);
        }

        private static void RequirePositive(int value, string name)
        {
            if (value <= 0)
            {
                throw new ArgumentException($"--{name} must be positive; got {value.ToString(CultureInfo.InvariantCulture)}.", nameof(value));
            }
        }

        // The column count only grows while scanning, so the tail columns kept for an earlier row are a
        // superset of the ones finally shown.
        private readonly record struct ColumnFilter(HashSet<int>? Wanted, int HeadColumns, int TailFrom)
        {
            internal bool Keeps(int column)
            {
                return Wanted?.Contains(column) ?? (column < HeadColumns || column >= TailFrom);
            }
        }

        private static ExcelColumnType TypeOf(IReadOnlyList<ExcelColumnSchema> schema, int column)
        {
            return column < schema.Count ? schema[column].Type : ExcelColumnType.StringColumn;
        }

        private static string?[] Materialize(Row row, bool isDate1904, IReadOnlyList<ExcelColumnSchema> schema, ColumnFilter filter)
        {
            string?[] values = new string?[row.ColumnCount];
            foreach (RowCell cell in row.Cells)
            {
                if (!filter.Keeps(cell.ColumnIndex))
                {
                    continue;
                }
                Cell value = cell.Value;
                values[cell.ColumnIndex] = FormatCell(in value, isDate1904, TypeOf(schema, cell.ColumnIndex));
            }
            return values;
        }

        private static string? FormatCell(in Cell cell, bool isDate1904, ExcelColumnType type)
        {
            switch (cell.Type)
            {
                case CellType.Empty:
                    return null;
                case CellType.Boolean:
                    if (!ExcelCellReaders.Bool(in cell, default, out bool flag))
                    {
                        return cell.GetString();
                    }
                    return flag ? "true" : "false";
                case CellType.Date:
                    if (!cell.TryGetDateTime(isDate1904, out DateTime date))
                    {
                        return cell.GetString();
                    }
                    return FormatDate(date, type);
                case CellType.Number or CellType.ExcelString when type == ExcelColumnType.Float64Column && cell.TryGetDouble(out double number):
                    return FormatFloat(number);
                default:
                    return cell.GetString();
            }
        }

        private static string FormatDate(DateTime date, ExcelColumnType type)
        {
            if (type != ExcelColumnType.TimestampColumn && date.TimeOfDay == TimeSpan.Zero)
            {
                return date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
            }
            return date.ToString("yyyy-MM-dd HH:mm:ss.FFFFFF", CultureInfo.InvariantCulture);
        }

        private static string FormatFloat(double number)
        {
            string text = number.ToString(CultureInfo.InvariantCulture);
            if (!double.IsFinite(number) || text.AsSpan().IndexOfAny('.', 'E') >= 0)
            {
                return text;
            }
            return text + ".0";
        }

        private static string PolarsType(ExcelColumnType type)
        {
            return type switch
            {
                ExcelColumnType.Int64Column => "i64",
                ExcelColumnType.Float64Column => "f64",
                ExcelColumnType.BoolColumn => "bool",
                ExcelColumnType.DateColumn => "date",
                ExcelColumnType.TimestampColumn => "datetime",
                _ => "str",
            };
        }

        private const string Ellipsis = "…";

        // ponytail: width is UTF-16 length, so wide East Asian glyphs and emoji misalign the borders.
        private static string Fit(string text, int maxWidth)
        {
            string flat = string.Create(text.Length, text, static (span, source) =>
            {
                for (int i = 0; i < span.Length; i++)
                {
                    span[i] = char.IsControl(source[i]) ? ' ' : source[i];
                }
            });
            if (flat.Length <= maxWidth)
            {
                return flat;
            }
            int cut = maxWidth - 1;
            if (cut > 0 && char.IsHighSurrogate(flat[cut - 1]))
            {
                cut--;
            }
            return string.Concat(flat.AsSpan(0, cut), Ellipsis);
        }

        private static List<int>? ResolveColumns(string? columns, IReadOnlyList<ExcelColumnSchema> schema)
        {
            if (columns is null)
            {
                return null;
            }
            List<int> resolved = [.. columns.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries).Select(part => ResolveColumn(part, schema))];
            if (resolved.Count == 0)
            {
                throw new ArgumentException("--columns lists no column.", nameof(columns));
            }
            return resolved;
        }

        private static int ResolveColumn(string part, IReadOnlyList<ExcelColumnSchema> schema)
        {
            for (int i = 0; i < schema.Count; i++)
            {
                if (string.Equals(ColumnName(schema, i), part, StringComparison.Ordinal))
                {
                    return i;
                }
            }
            if (!int.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out int index))
            {
                throw new ArgumentException($"--columns: no column named '{part}'.", nameof(part));
            }
            if (index >= schema.Count)
            {
                throw new ArgumentException($"--columns: index {part} is out of range; the sheet has {schema.Count.ToString(CultureInfo.InvariantCulture)} column(s).", nameof(part));
            }
            return index;
        }

        private static List<int> ShownColumns(List<int> selected, int maxColumns)
        {
            if (selected.Count <= maxColumns)
            {
                return selected;
            }
            int headColumns = (maxColumns + 1) / 2;
            int tailColumns = maxColumns - headColumns;
            return [.. selected[..headColumns], -1, .. selected[^tailColumns..]];
        }

        private static void WriteTable(TextWriter stdout, IReadOnlyList<ExcelColumnSchema> schema, string total, List<int> selected, List<string?[]?> body, int maxColumns, int maxWidth)
        {
            List<int> shown = ShownColumns(selected, maxColumns);
            List<string[]> lines = [];
            lines.Add([.. shown.Select(c => c < 0 ? Ellipsis : Fit(ColumnName(schema, c), maxWidth))]);
            lines.Add([.. shown.Select(static _ => "---")]);
            lines.Add([.. shown.Select(c => c < 0 ? Ellipsis : PolarsType(TypeOf(schema, c)))]);
            foreach (string?[]? row in body)
            {
                lines.Add([.. shown.Select(c => BodyCell(row, c, maxWidth))]);
            }

            int[] widths = new int[shown.Count];
            foreach (string[] line in lines)
            {
                for (int i = 0; i < widths.Length; i++)
                {
                    widths[i] = Math.Max(widths[i], line[i].Length);
                }
            }

            stdout.WriteLine($"shape: ({total}, {Underscored(selected.Count)})");
            WriteBorder(stdout, widths, '┌', '─', '┬', '┐');
            for (int i = 0; i < 3; i++)
            {
                WriteCells(stdout, widths, lines[i]);
            }
            WriteBorder(stdout, widths, '╞', '═', '╪', '╡');
            for (int i = 3; i < lines.Count; i++)
            {
                WriteCells(stdout, widths, lines[i]);
            }
            WriteBorder(stdout, widths, '└', '─', '┴', '┘');
        }

        private static string ColumnName(IReadOnlyList<ExcelColumnSchema> schema, int column)
        {
            string? name = column < schema.Count ? schema[column].Name : null;
            return name ?? $"column_{(column + 1).ToString(CultureInfo.InvariantCulture)}";
        }

        private static string BodyCell(string?[]? row, int column, int maxWidth)
        {
            if (row is null || column < 0)
            {
                return Ellipsis;
            }
            string? value = column < row.Length ? row[column] : null;
            return Fit(value ?? "null", maxWidth);
        }

        private static string Underscored(long value)
        {
            return value.ToString("N0", CultureInfo.InvariantCulture).Replace(',', '_');
        }

        private static void WriteBorder(TextWriter stdout, int[] widths, char left, char fill, char join, char right)
        {
            stdout.Write(left);
            for (int i = 0; i < widths.Length; i++)
            {
                if (i > 0)
                {
                    stdout.Write(join);
                }
                stdout.Write(new string(fill, widths[i] + 2));
            }
            stdout.WriteLine(right);
        }

        private static void WriteCells(TextWriter stdout, int[] widths, string[] cells)
        {
            stdout.Write("│ ");
            for (int i = 0; i < widths.Length; i++)
            {
                if (i > 0)
                {
                    stdout.Write(" ┆ ");
                }
                stdout.Write(cells[i].PadRight(widths[i]));
            }
            stdout.WriteLine(" │");
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
                stderr.WriteLine(UserMessage(exception));
                return 1;
            }
        }

        private static string UserMessage(Exception exception)
        {
            if (exception is not ArgumentException { ParamName: string name })
            {
                return exception.Message;
            }
            return exception.Message.Replace($" (Parameter '{name}')", string.Empty, StringComparison.Ordinal);
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
            CsvReaderOptions csv = inputDelimiter is char delimiter
                ? CsvReaderOptions.Default with { Delimiter = AsciiByte(delimiter, "input-delimiter") }
                : CsvReaderOptions.Default with { SniffDialect = true };
            ExcelReaderOptions options = new() { Csv = csv };
            if (password is not null)
            {
                options = options with { Password = password };
            }
            IExcelWorkbook workbook = Excel.Open(path, IsCsv(path) ? ExcelFileFormat.Csv : ExcelFileFormat.Unknown, options);
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

        private static bool IsCsv(string path)
        {
            return string.Equals(Path.GetExtension(path), ".csv", StringComparison.OrdinalIgnoreCase);
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
