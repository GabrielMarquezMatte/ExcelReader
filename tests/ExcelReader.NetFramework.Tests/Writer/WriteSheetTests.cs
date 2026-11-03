using ExcelReader.NetFramework.Parser;
using ExcelReader.NetFramework.Reader;
using ExcelReader.NetFramework.Writer;

namespace ExcelReader.NetFramework.Tests.Writer
{
    public sealed class WriteRow
    {
        [ExcelColumn("Nome")]
        public string? Name { get; set; }

        public long Id { get; set; }

        public int Qty { get; set; }

        public short Small { get; set; }

        public byte Tiny { get; set; }

        public double Price { get; set; }

        public float Ratio { get; set; }

        public decimal Amount { get; set; }

        public bool Active { get; set; }

        public DateTime When { get; set; }

        public TimeSpan At { get; set; }

        public int? MaybeQty { get; set; }

        public DateTime? MaybeWhen { get; set; }

        [ExcelIgnore]
        public string? Secret { get; set; }
    }

    public sealed class GuidRow
    {
        public string? Name { get; set; }

        public Guid Key { get; set; }
    }

    internal struct StructRow
    {
        public string? Label { get; set; }

        public int Count { get; set; }
    }

    public sealed class TimeRow
    {
        public TimeSpan At { get; set; }
    }

    public sealed class WriteSheetTests
    {
        private static readonly string[] Header = { "Nome", "Id", "Qty", "Small", "Tiny", "Price", "Ratio", "Amount", "Active", "When", "At", "MaybeQty", "MaybeWhen" };
        private static readonly string[] SheetNames = { "Resumo", "Vendas" };

        private static List<WriteRow> Rows()
        {
            return new List<WriteRow>
            {
                new WriteRow
                {
                    Name = "ação 😀", Id = 9_000_000_000L, Qty = 42, Small = -1234, Tiny = 200, Price = 12.5, Ratio = 0.25f, Amount = 1234.5m, Active = true,
                    When = new DateTime(2026, 3, 15, 10, 20, 30), At = new TimeSpan(8, 45, 0), MaybeQty = 7, MaybeWhen = new DateTime(2025, 12, 31, 23, 59, 59), Secret = "hidden",
                },
                new WriteRow
                {
                    Name = "plain", Id = -1, Qty = 0, Small = 1, Tiny = 0, Price = -0.75, Ratio = 1.5f, Amount = -2m, Active = false,
                    When = new DateTime(1999, 1, 1), At = new TimeSpan(23, 0, 1), MaybeQty = -3, MaybeWhen = new DateTime(2000, 2, 29, 12, 0, 0),
                },
                new WriteRow
                {
                    Name = null, Id = 3, Qty = 3, Small = 3, Tiny = 3, Price = 3, Ratio = 3, Amount = 3, Active = true,
                    When = new DateTime(2026, 1, 1, 0, 0, 1), At = TimeSpan.Zero, MaybeQty = null, MaybeWhen = null,
                },
            };
        }

        private static byte[] WriteVendas(ExcelFormat format, IEnumerable<WriteRow> rows, ExcelTableOptions? table = null)
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(format);
            writer.WriteSheet("Vendas", rows, table);
            return writer.CloseToArray();
        }

        private static ExcelWorkbook Open(byte[] bytes, ExcelFormat format)
        {
            return ExcelWorkbook.Open(bytes, new ExcelOpenOptions { Format = format });
        }

        private static void AssertRowsEqual(List<WriteRow> expected, WriteRow[] actual)
        {
            Assert.Equal(expected.Count, actual.Length);
            for (int i = 0; i < expected.Count; i++)
            {
                WriteRow e = expected[i];
                WriteRow a = actual[i];
                Assert.Equal(e.Name, a.Name);
                Assert.Equal(e.Id, a.Id);
                Assert.Equal(e.Qty, a.Qty);
                Assert.Equal(e.Small, a.Small);
                Assert.Equal(e.Tiny, a.Tiny);
                Assert.Equal(e.Price, a.Price);
                Assert.Equal(e.Ratio, a.Ratio);
                Assert.Equal((double)e.Amount, (double)a.Amount);
                Assert.Equal(e.Active, a.Active);
                Assert.Equal(e.When, a.When);
                Assert.Equal(e.At, a.At);
                Assert.Equal(e.MaybeQty, a.MaybeQty);
                Assert.Equal(e.MaybeWhen, a.MaybeWhen);
                Assert.Null(a.Secret);
            }
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx)]
        [InlineData(ExcelFormat.Xlsb)]
        [InlineData(ExcelFormat.Xls)]
        [InlineData(ExcelFormat.Csv)]
        public void TypedRoundTrip(ExcelFormat format)
        {
            List<WriteRow> rows = Rows();
            using ExcelWorkbook workbook = Open(WriteVendas(format, rows), format);

            AssertRowsEqual(rows, workbook.Sheet(0).Read<WriteRow>());
        }

        [Fact]
        public void HeaderUsesFirstColumnName()
        {
            using ExcelWorkbook workbook = Open(WriteVendas(ExcelFormat.Xlsx, Rows()), ExcelFormat.Xlsx);

            ExcelRow header = workbook.Sheet(0).ReadRows().First();

            Assert.Equal(Header, header.Cells.Select(c => c.Text).ToArray());
        }

        [Fact]
        public void EmptyListWritesHeaderOnly()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("E", Array.Empty<WriteRow>());
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);

            Assert.Single(workbook.Sheet("E").ReadRows());
            Assert.Empty(workbook.Sheet("E").Read<WriteRow>());
        }

        [Fact]
        public void MixesWithCellWriting()
        {
            List<WriteRow> rows = Rows();
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.StartSheet("Resumo");
            writer.StartRow();
            writer.Write("total");
            writer.Write(3L);
            writer.EndRow();
            writer.EndSheet();
            writer.WriteSheet("Vendas", rows);
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);

            Assert.Equal(SheetNames, workbook.Sheets.Select(s => s.Name).ToArray());
            Assert.Equal("total|3", string.Join("|", workbook.Sheet("Resumo").ReadRows().Single().Cells.Select(c => c.Text)));
            AssertRowsEqual(rows, workbook.Sheet("Vendas").Read<WriteRow>());
        }

        [Fact]
        public void WriteSheetWhileSheetOpenThrows()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.StartSheet("x");

            Assert.Throws<InvalidOperationException>(() => writer.WriteSheet("y", Rows()));

            writer.EndSheet();
            writer.WriteSheet("y", Rows());
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);
            Assert.Equal(2, workbook.Sheets.Count);
        }

        [Fact]
        public void UnsupportedTypeThrowsBeforeWriting()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            GuidRow[] rows = { new GuidRow { Name = "a", Key = Guid.NewGuid() } };

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => writer.WriteSheet("Bad", rows));

            Assert.Contains("GuidRow.Key", exception.Message, StringComparison.Ordinal);
            writer.StartSheet("Ok");
            writer.EndSheet();
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);
            Assert.Equal("Ok", workbook.Sheets.Single().Name);
        }

        [Theory]
        [InlineData(ExcelFormat.Xlsx)]
        [InlineData(ExcelFormat.Xlsb)]
        public void TableOptions(ExcelFormat format)
        {
            List<WriteRow> rows = Rows();
            using ExcelWorkbook workbook = Open(WriteVendas(format, rows, new ExcelTableOptions { Name = "TabelaVendas" }), format);

            ExcelTableInfo table = Assert.Single(workbook.Tables);
            Assert.Equal("TabelaVendas", table.Name);
            Assert.Equal(0, table.SheetIndex);
            Assert.Equal(1, table.HeaderRowCount);
            Assert.Equal(rows.Count, table.LastRow - table.FirstRow);
            AssertRowsEqual(rows, workbook.Table("TabelaVendas").Read<WriteRow>());
        }

        [Fact]
        public void TableOnCsvThrows()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Csv);

            Assert.Throws<ExcelReaderException>(() => writer.WriteSheet("Vendas", Rows(), new ExcelTableOptions { Name = "TabelaVendas" }));
        }

        [Fact]
        public void StructRowsWrite()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);
            writer.WriteSheet("S", new[] { new StructRow { Label = "a", Count = 1 }, new StructRow { Label = "b", Count = 2 } });
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);

            Assert.Equal("Label,Count|a,1|b,2", string.Join("|", workbook.Sheet(0).ReadRows().Select(r => string.Join(",", r.Cells.Select(c => c.Text)))));
        }

        [Fact]
        public void TimeOutsideOneDayThrowsBeforeWriting()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);

            Assert.Throws<ArgumentOutOfRangeException>(() => writer.WriteSheet("T", new[] { new TimeRow { At = TimeSpan.FromHours(25) } }));

            writer.WriteSheet("T", new[] { new TimeRow { At = TimeSpan.FromHours(1) } });
            using ExcelWorkbook workbook = Open(writer.CloseToArray(), ExcelFormat.Xlsx);
            Assert.Equal(TimeSpan.FromHours(1), workbook.Sheet("T").Read<TimeRow>().Single().At);
        }

        [Fact]
        public void TableWithoutNameThrows()
        {
            using ExcelWorkbookWriter writer = ExcelWorkbookWriter.CreateInMemory(ExcelFormat.Xlsx);

            Assert.Throws<ArgumentException>(() => writer.WriteSheet("Vendas", Rows(), new ExcelTableOptions()));
        }
    }
}
