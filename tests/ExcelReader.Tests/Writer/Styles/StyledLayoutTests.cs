using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Csv;
using ExcelReader.Core.Writer.Internal;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Writer.Styles
{
    public class StyledLayoutTests
    {
        public static TheoryData<string> Formats => ["xlsx", "xlsb", "xls"];

        public static TheoryData<string> TableFormats => ["xlsx", "xlsb"];

        private static readonly CellStyle HeaderStyle = new() { Bold = true, FillColor = ExcelColor.FromRgb(0xFFFF00) };

        private static CancellationToken Ct => TestContext.Current.CancellationToken;

        public sealed class Payment
        {
            public string Name { get; init; } = "";

            [ExcelStyle(Bold = true)]
            public double Amount { get; init; }
        }

        public sealed class Plain
        {
            public string Name { get; init; } = "";

            public double Amount { get; init; }
        }

        public sealed class TooBig
        {
            [ExcelStyle(FontSize = 500)]
            public double Amount { get; init; }
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

        private static readonly Payment[] Payments = [new() { Name = "a", Amount = 1.5 }, new() { Name = "b", Amount = 2 }];
        private static readonly Plain[] Plains = [new() { Name = "a", Amount = 1.5 }, new() { Name = "b", Amount = 2 }];

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task AttributeStyleAppliesToDataCellsOnly(string format)
        {
            byte[] file = await WriteAsync(format, t => t.WriteAsync(Payments, ExcelRecordLayout.FromAttributes<Payment>()));
            Assert.False(Probe(format, file, 0, 1).Bold);
            Assert.True(Probe(format, file, 1, 1).Bold);
            Assert.True(Probe(format, file, 2, 1).Bold);
            Assert.False(Probe(format, file, 1, 0).Bold);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task AttributeStyleBeatsColumnStyle(string format)
        {
            byte[] file = await WriteAsync(format, t =>
            {
                int italic = t.AddStyle(new CellStyle { Italic = true });
                return t.WriteAsync(Payments, ExcelRecordLayout.FromAttributes<Payment>(), styledColumn: 1, columnStyle: italic);
            });
            ProbedStyle data = Probe(format, file, 1, 1);
            Assert.True(data.Bold);
            Assert.False(data.Italic);
            Assert.True(Probe(format, file, 0, 1).Italic);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task HeaderStyleAppliesToEveryHeaderCell(string format)
        {
            ExcelRecordLayout<Plain> layout = ExcelRecordLayout.FromAttributes<Plain>().WithHeaderStyle(HeaderStyle);
            byte[] file = await WriteAsync(format, t => t.WriteAsync(Plains, layout));
            for (int column = 0; column < 2; column++)
            {
                ProbedStyle header = Probe(format, file, 0, column);
                Assert.True(header.Bold);
                Assert.Equal(0xFFFF00, header.FillRgb);
                ProbedStyle data = Probe(format, file, 1, column);
                Assert.False(data.Bold);
                Assert.Null(data.FillRgb);
            }
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task MappedLayoutValueDependentStyle(string format)
        {
            Balance[] balances = [new() { Value = 5 }, new() { Value = -3 }, new() { Value = 7 }];
            byte[] file = await WriteAsync(format, t =>
            {
                Balance.Red = t.AddStyle(new CellStyle { FontColor = ExcelColor.FromRgb(0xFF0000) });
                return t.WriteAsync(balances, ExcelRecordLayout.Generated<Balance>());
            });
            Assert.Null(Probe(format, file, 1, 0).FontRgb);
            Assert.Equal(0xFF0000, Probe(format, file, 2, 0).FontRgb);
            Assert.Null(Probe(format, file, 3, 0).FontRgb);
        }

        [Theory]
        [MemberData(nameof(Formats))]
        public async Task SameLayoutOnTwoSheetsReusesStyleIds(string format)
        {
            ExcelRecordLayout<Payment> layout = ExcelRecordLayout.FromAttributes<Payment>().WithHeaderStyle(HeaderStyle);
            byte[] file = await WriteAsync(format, async t =>
            {
                await t.WriteAsync(Payments, layout);
                int count = t.StyleCount;
                await t.WriteAsync(Payments, layout);
                Assert.Equal(count, t.StyleCount);
            });
            Assert.Equal(StyleIndex(format, file, 0, 1, 1), StyleIndex(format, file, 1, 1, 1));
            Assert.Equal(StyleIndex(format, file, 0, 0, 0), StyleIndex(format, file, 1, 0, 0));
        }

        [Theory]
        [MemberData(nameof(TableFormats))]
        public async Task TableHeaderTakesHeaderStyle(string format)
        {
            ExcelRecordLayout<Plain> layout = ExcelRecordLayout.FromAttributes<Plain>().WithHeaderStyle(HeaderStyle);
            byte[] file = await WriteAsync(format, t => t.WriteTableAsync(Plains, layout, new ExcelTableOptions { FirstColumn = 2 }));
            Assert.True(Probe(format, file, 0, 2).Bold);
            Assert.Equal(0xFFFF00, Probe(format, file, 0, 3).FillRgb);
            Assert.False(Probe(format, file, 1, 2).Bold);
            Assert.Null(Probe(format, file, 1, 3).FillRgb);
        }

        [Fact]
        public void InvalidAttributeThrowsFromFromAttributes()
        {
            Assert.Throws<ArgumentException>(ExcelRecordLayout.FromAttributes<TooBig>);
        }

        [Fact]
        public void WithHeaderStyleValidates()
        {
            ExcelRecordLayout<Plain> layout = ExcelRecordLayout.FromAttributes<Plain>();
            Assert.Throws<ArgumentException>(() => layout.WithHeaderStyle(new CellStyle { FontSize = 0.5 }));
        }

        [Fact]
        public void UnstyledLayoutKeepsTheCachedPlan()
        {
            using var ms = new MemoryStream();
            using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(ms, leaveOpen: true);
            XlsxSheetWriter sheet = workbook.AddSheet("S1");
            ExcelRecordLayout<Plain> layout = ExcelRecordLayout.FromAttributes<Plain>();
            int styles = workbook.StyleCount;
            BoundLayout<XlsxRowWriter, Plain> bound = LayoutBinding.Bind(sheet, layout);
            Assert.Same(layout.Columns<XlsxRowWriter>().WriteRow, bound.WriteRow);
            Assert.Same(layout.Columns<XlsxRowWriter>().Headers, bound.Headers);
            Assert.Equal(-1, bound.HeaderStyleId);
            Assert.Null(bound.ColumnStyleIds);
            Assert.Equal(styles, workbook.StyleCount);
        }

        [Fact]
        public async Task StyledLayoutOnExternalSheetThrows()
        {
            var sheet = new ExternalSheet();
            await Assert.ThrowsAsync<NotSupportedException>(
                () => sheet.WriteRecordsAsync(Payments, ExcelRecordLayout.FromAttributes<Payment>(), Ct).AsTask());
            await Assert.ThrowsAsync<NotSupportedException>(
                () => sheet.WriteRecordsAsync(Plains, ExcelRecordLayout.FromAttributes<Plain>().WithHeaderStyle(HeaderStyle), Ct).AsTask());
            Assert.Equal(0, sheet.RowsStarted);
        }

        [Fact]
        public async Task CsvAcceptsStyledLayouts()
        {
            string styled = await CsvAsync(Payments, ExcelRecordLayout.FromAttributes<Payment>().WithHeaderStyle(HeaderStyle));
            string plain = await CsvAsync(Plains, ExcelRecordLayout.FromAttributes<Plain>());
            Assert.Equal(plain, styled);
            Assert.StartsWith("Name,Amount", styled, StringComparison.Ordinal);
        }

        private static async Task<string> CsvAsync<T>(IEnumerable<T> records, ExcelRecordLayout<T> layout)
        {
            using var ms = new MemoryStream();
            await using (CsvWorkbookWriter workbook = CsvWorkbookWriter.Create(ms, leaveOpen: true))
            {
                CsvSheetWriter sheet = workbook.AddSheet("S1");
                await sheet.WriteRecordsAsync(records, layout, Ct);
                await sheet.EndAsync(Ct);
                await workbook.EndAsync(Ct);
            }
            return Encoding.UTF8.GetString(ms.ToArray());
        }

        private static async Task<byte[]> WriteAsync(string format, Func<ITarget, Task> script)
        {
            var ms = new MemoryStream();
            ITarget target = Open(format, ms);
            await script(target);
            await target.FinishAsync();
            return ms.ToArray();
        }

        private static ITarget Open(string format, MemoryStream ms)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return new Target<XlsxWorkbookWriter, XlsxSheetWriter, XlsxRowWriter>(XlsxWorkbookWriter.Create(ms, leaveOpen: true), static w => w.StyleCount);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return new Target<XlsbWorkbookWriter, XlsbSheetWriter, XlsbRowWriter>(XlsbWorkbookWriter.Create(ms, leaveOpen: true), static w => w.StyleCount);
            }
            return new Target<XlsWorkbookWriter, XlsSheetWriter, XlsRowWriter>(XlsWorkbookWriter.Create(ms, leaveOpen: true), static w => w.StyleCount);
        }

        private static ProbedStyle Probe(string format, byte[] file, int row, int column)
        {
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsx(file, row, column);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return StyleProbe.Xlsb(file, row, column);
            }
            return StyleProbe.Xls(file, row, column);
        }

        private static int StyleIndex(string format, byte[] file, int sheet, int row, int column)
        {
            using IExcelWorkbook reader = OpenReader(format, file);
            using var rows = reader.SheetAt(sheet).GetEnumerator();
            for (int i = 0; i <= row; i++)
            {
                Assert.True(rows.MoveNext());
            }
            return rows.Current[column].StyleIndex;
        }

        private static IExcelWorkbook OpenReader(string format, byte[] file)
        {
            var ms = new MemoryStream(file);
            if (string.Equals(format, "xlsx", StringComparison.Ordinal))
            {
                return Excel.FromXlsx(ms);
            }
            if (string.Equals(format, "xlsb", StringComparison.Ordinal))
            {
                return Excel.FromXlsb(ms);
            }
            return Excel.FromXls(ms);
        }

        private interface ITarget
        {
            int StyleCount { get; }

            int AddStyle(CellStyle style);

            Task WriteAsync<T>(IEnumerable<T> records, ExcelRecordLayout<T> layout, int styledColumn = -1, int columnStyle = 0);

            Task WriteTableAsync<T>(IEnumerable<T> records, ExcelRecordLayout<T> layout, ExcelTableOptions options);

            Task FinishAsync();
        }

        private sealed class Target<TWorkbook, TSheet, TRow>(TWorkbook workbook, Func<TWorkbook, int> styleCount) : ITarget
            where TWorkbook : IWorkbookWriter<TSheet>
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            private int _sheets;

            public int StyleCount => styleCount(workbook);

            public int AddStyle(CellStyle style)
            {
                return workbook.AddStyle(style);
            }

            public async Task WriteAsync<T>(IEnumerable<T> records, ExcelRecordLayout<T> layout, int styledColumn = -1, int columnStyle = 0)
            {
                ISheetWriter<TRow> sheet = workbook.AddSheet($"S{++_sheets}");
                if (styledColumn >= 0)
                {
                    sheet.SetColumnStyle(styledColumn, columnStyle);
                }
                await sheet.WriteRecordsAsync(records, layout, Ct);
                await sheet.EndAsync(Ct);
            }

            public async Task WriteTableAsync<T>(IEnumerable<T> records, ExcelRecordLayout<T> layout, ExcelTableOptions options)
            {
                var sheet = (ITableSheetWriter<TRow>)workbook.AddSheet($"S{++_sheets}");
                await sheet.WriteTableAsync(records, "Payments", layout, options, Ct);
                await sheet.EndAsync(Ct);
            }

            public async Task FinishAsync()
            {
                await workbook.EndAsync(Ct);
                await workbook.DisposeAsync();
            }
        }

        private sealed class ExternalSheet : ISheetWriter<XlsxRowWriter>
        {
            internal int RowsStarted { get; private set; }

            public XlsxRowWriter StartRow()
            {
                RowsStarted++;
                throw new InvalidOperationException("No rows on this sheet.");
            }

            public XlsxRowWriter StartRow(int styleId)
            {
                return StartRow();
            }

            public ValueTask<XlsxRowWriter> StartRowAsync(CancellationToken ct = default)
            {
                return ValueTask.FromResult(StartRow());
            }

            public ValueTask<XlsxRowWriter> StartRowAsync(int styleId, CancellationToken ct = default)
            {
                return ValueTask.FromResult(StartRow());
            }

            public void SetColumnStyle(int columnIndex, int styleId) { }

            public void SetColumnWidth(int columnIndex, double width) { }

            public void End() { }

            public ValueTask EndAsync(CancellationToken ct = default)
            {
                return ValueTask.CompletedTask;
            }

            public void Dispose() { }

            public ValueTask DisposeAsync()
            {
                return ValueTask.CompletedTask;
            }
        }
    }
}
