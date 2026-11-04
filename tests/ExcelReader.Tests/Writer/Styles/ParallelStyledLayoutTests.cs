using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class ParallelStyledLayoutTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb"];

        private static readonly CellStyle HeaderStyle = new() { Bold = true, FillColor = ExcelColor.FromRgb(0xFFFF00) };

        private static readonly ParallelWriteOptions Parallel = new() { DegreeOfParallelism = 4, BatchRows = 512 };

        private static CancellationToken Ct => TestContext.Current.CancellationToken;

        public sealed class Entry
        {
            public string Name { get; init; } = "";

            [ExcelStyle(Bold = true, NumberFormat = "0.00")]
            public double Amount { get; init; }

            [ExcelStyle(Bold = true)]
            public DateTime When { get; init; }
        }

        public sealed class Balance : IExcelRecordMap<Balance>
        {
            internal static int Red;

            public double Value { get; init; }

            public static void ConfigureExcelRecordMap<TRow>(ExcelRecordMapBuilder<Balance, TRow> builder)
                where TRow : IRowWriter
            {
                builder.Column("Value", static (row, record) =>
                {
                    if (record.Value < 0)
                    {
                        row.Style(Red);
                    }
                    row.Write(record.Value);
                });
            }
        }

        private static readonly Entry[] Entries = [.. Enumerable.Range(0, 10_000).Select(i => new Entry
        {
            Name = $"n{i}",
            Amount = i * 1.5,
            When = new DateTime(2026, 1, 1).AddDays(i % 365),
        })];

        private static readonly Balance[] Balances = [.. Enumerable.Range(0, 10_000).Select(i => new Balance { Value = i % 3 == 0 ? -i : i })];

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task ParallelStyledWriteMatchesSerial(string format)
        {
            ExcelRecordLayout<Entry> layout = ExcelRecordLayout.FromAttributes<Entry>().WithHeaderStyle(HeaderStyle);
            byte[] serial = await WriteAsync(format, parallel: false, layout, table: false);
            byte[] parallel = await WriteAsync(format, parallel: true, layout, table: false);
            Assert.Equal(StyleIndices(format, serial, 0), StyleIndices(format, parallel, 0));
            Assert.Equal(StyleIndices(format, serial, 1), StyleIndices(format, parallel, 1));
            Assert.True(Probe(format, parallel, 0, 0).Bold);
            Assert.True(Probe(format, parallel, Entries.Length, 1).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task ParallelTableHeaderTakesHeaderStyle(string format)
        {
            ExcelRecordLayout<Entry> layout = ExcelRecordLayout.FromAttributes<Entry>().WithHeaderStyle(HeaderStyle);
            byte[] file = await WriteAsync(format, parallel: true, layout, table: true);
            ProbedStyle header = Probe(format, file, 0, 1);
            Assert.True(header.Bold);
            Assert.Equal(0xFFFF00, header.FillRgb);
            Assert.Null(Probe(format, file, 1, 0).FillRgb);
            Assert.True(Probe(format, file, 1, 2).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task ParallelStyledDateColumnReadsAsDates(string format)
        {
            byte[] file = await WriteAsync(format, parallel: true, ExcelRecordLayout.FromAttributes<Entry>(), table: false);
            using IExcelWorkbook reader = OpenReader(format, file);
            int rows = 0;
            foreach (Row row in reader.SheetAt(0))
            {
                if (rows > 0)
                {
                    Assert.Equal(CellType.Date, row[2].Type);
                }
                rows++;
            }
            Assert.Equal(Entries.Length + 1, rows);
            Assert.True(Probe(format, file, Entries.Length, 2).Bold);
        }

        private static async Task<byte[]> WriteAsync(string format, bool parallel, ExcelRecordLayout<Entry> layout, bool table)
        {
            var ms = new MemoryStream();
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                await using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
                Balance.Red = workbook.AddStyle(new CellStyle { FontColor = ExcelColor.FromRgb(0xFF0000) });
                XlsxSheetWriter first = workbook.AddSheet("Entries");
                await WriteEntriesAsync(first, parallel, layout, table);
                XlsxSheetWriter second = workbook.AddSheet("Balances");
                await WriteBalancesAsync(second, parallel);
                await workbook.EndAsync(Ct);
                return ms.ToArray();
            }
            await using XlsbWorkbookWriter book = XlsbWorkbookWriter.Create(ms, leaveOpen: true);
            Balance.Red = book.AddStyle(new CellStyle { FontColor = ExcelColor.FromRgb(0xFF0000) });
            XlsbSheetWriter entries = book.AddSheet("Entries");
            await WriteEntriesAsync(entries, parallel, layout, table);
            XlsbSheetWriter balances = book.AddSheet("Balances");
            await WriteBalancesAsync(balances, parallel);
            await book.EndAsync(Ct);
            return ms.ToArray();
        }

        private static async Task WriteEntriesAsync(XlsxSheetWriter sheet, bool parallel, ExcelRecordLayout<Entry> layout, bool table)
        {
            if (table)
            {
                await sheet.WriteTableParallelAsync(Entries, "Entries", layout, tableOptions: null, Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            if (parallel)
            {
                await sheet.WriteRecordsParallelAsync(Entries, layout, Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            await sheet.WriteRecordsAsync(Entries, layout, Ct);
            await sheet.EndAsync(Ct);
        }

        private static async Task WriteEntriesAsync(XlsbSheetWriter sheet, bool parallel, ExcelRecordLayout<Entry> layout, bool table)
        {
            if (table)
            {
                await sheet.WriteTableParallelAsync(Entries, "Entries", layout, tableOptions: null, Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            if (parallel)
            {
                await sheet.WriteRecordsParallelAsync(Entries, layout, Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            await sheet.WriteRecordsAsync(Entries, layout, Ct);
            await sheet.EndAsync(Ct);
        }

        private static async Task WriteBalancesAsync(XlsxSheetWriter sheet, bool parallel)
        {
            if (parallel)
            {
                await sheet.WriteRecordsParallelAsync(Balances, ExcelRecordLayout.Generated<Balance>(), Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            await sheet.WriteRecordsAsync(Balances, ExcelRecordLayout.Generated<Balance>(), Ct);
            await sheet.EndAsync(Ct);
        }

        private static async Task WriteBalancesAsync(XlsbSheetWriter sheet, bool parallel)
        {
            if (parallel)
            {
                await sheet.WriteRecordsParallelAsync(Balances, ExcelRecordLayout.Generated<Balance>(), Parallel, Ct);
                await sheet.EndAsync(Ct);
                return;
            }
            await sheet.WriteRecordsAsync(Balances, ExcelRecordLayout.Generated<Balance>(), Ct);
            await sheet.EndAsync(Ct);
        }

        private static List<int> StyleIndices(string format, byte[] file, int sheet)
        {
            using IExcelWorkbook reader = OpenReader(format, file);
            var indices = new List<int>();
            foreach (Row row in reader.SheetAt(sheet))
            {
                for (int column = 0; column < row.ColumnCount; column++)
                {
                    indices.Add(row[column].StyleIndex);
                }
            }
            return indices;
        }

        private static ProbedStyle Probe(string format, byte[] file, int row, int column)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsx(file, row, column);
            }
            return StyleProbe.Xlsb(file, row, column);
        }

        private static IExcelWorkbook OpenReader(string format, byte[] file)
        {
            var ms = new MemoryStream(file);
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return Excel.FromXlsx(ms);
            }
            return Excel.FromXlsb(ms);
        }
    }
}
