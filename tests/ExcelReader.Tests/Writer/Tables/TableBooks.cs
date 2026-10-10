using System.IO.Compression;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Tables
{
    internal interface ITableBook
    {
        void AddSheet(string name);

        void BeginTable(string name, string[] columns, ExcelTableOptions? options = null);

        ValueTask BeginTableAsync(string name, string[] columns, ExcelTableOptions? options = null);

        void Row(params object?[] values);

        void EndTable();

        void EndSheet();

        byte[] Finish();

        ValueTask<byte[]> FinishAsync();
    }

    internal sealed class TableBook<TSheet, TRow>(IWorkbookWriter<TSheet> workbook, MemoryStream stream) : ITableBook
        where TSheet : class, ITableSheetWriter<TRow>
        where TRow : IRowWriter
    {
        private TSheet? _sheet;

        private TSheet Sheet
        {
            get
            {
                return _sheet ?? throw new InvalidOperationException("No sheet added.");
            }
        }

        public void AddSheet(string name)
        {
            _sheet = workbook.AddSheet(name);
        }

        public void BeginTable(string name, string[] columns, ExcelTableOptions? options = null)
        {
            Sheet.BeginTable(name, columns, options);
        }

        public ValueTask BeginTableAsync(string name, string[] columns, ExcelTableOptions? options = null)
        {
            return Sheet.BeginTableAsync(name, columns, options, TestContext.Current.CancellationToken);
        }

        public void Row(params object?[] values)
        {
            using TRow row = Sheet.StartRow();
            foreach (object? value in values)
            {
                WriteValue(row, value);
            }
        }

        public void EndTable()
        {
            Sheet.EndTable();
        }

        public void EndSheet()
        {
            Sheet.End();
            _sheet = null;
        }

        public byte[] Finish()
        {
            workbook.End();
            workbook.Dispose();
            return stream.ToArray();
        }

        public async ValueTask<byte[]> FinishAsync()
        {
            CancellationToken ct = TestContext.Current.CancellationToken;
            if (_sheet is not null)
            {
                await _sheet.EndAsync(ct);
                _sheet = null;
            }
            await workbook.EndAsync(ct);
            await workbook.DisposeAsync();
            return stream.ToArray();
        }

        private static void WriteValue(TRow row, object? value)
        {
            switch (value)
            {
                case null:
                    row.Skip();
                    break;
                case string text:
                    row.Write(text);
                    break;
                case int number:
                    row.Write(number);
                    break;
                case double real:
                    row.Write(real);
                    break;
                default:
                    throw new ArgumentException($"Unsupported cell value {value.GetType()}.", nameof(value));
            }
        }
    }

    public static class TableBooks
    {
        public static TheoryData<string> Formats { get; } = new() { "xlsx", "xlsb" };

        internal static ITableBook Create(string format)
        {
            MemoryStream stream = new();
            return format switch
            {
                "xlsx" => new TableBook<XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(stream, leaveOpen: true), stream),
                "xlsb" => new TableBook<XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(stream, leaveOpen: true), stream),
                _ => throw new ArgumentOutOfRangeException(nameof(format)),
            };
        }

        internal static string Extension(string format)
        {
            return string.Equals(format, "xlsx", StringComparison.Ordinal) ? "xml" : "bin";
        }

        internal static byte[] Entry(byte[] package, string name)
        {
            using ZipArchive zip = new(new MemoryStream(package), ZipArchiveMode.Read);
            ZipArchiveEntry entry = zip.GetEntry(name) ?? throw new InvalidOperationException($"Package has no entry '{name}'.");
            using Stream content = entry.Open();
            using MemoryStream copy = new();
            content.CopyTo(copy);
            return copy.ToArray();
        }

        internal static IReadOnlyList<string> EntryNames(byte[] package)
        {
            using ZipArchive zip = new(new MemoryStream(package), ZipArchiveMode.Read);
            return [.. zip.Entries.Select(static e => e.FullName)];
        }

        internal static Dictionary<string, byte[]> Entries(byte[] package)
        {
            Dictionary<string, byte[]> entries = new(StringComparer.Ordinal);
            foreach (string name in EntryNames(package))
            {
                entries[name] = Entry(package, name);
            }
            return entries;
        }
    }
}
