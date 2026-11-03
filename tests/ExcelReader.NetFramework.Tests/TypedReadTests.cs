using System.Text;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class TypedReadTests
    {
        private const int RealDataRows = 100;
        private const string QtyCsv = "Name,Qty\r\na,1\r\nb,\r\n";
        private static readonly int?[] OneThenNull = { 1, null };
        private static readonly int[] OneThenZero = { 1, 0 };
        private static readonly string?[] NullThenB = { null, "b" };

        [Theory]
        [InlineData("RealExcel.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void ReadsTypedRows(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path(file));
            Row[] rows = workbook.Sheet(0).Read<Row>();
            Assert.Equal(RealDataRows, rows.Length);
            Assert.Equal("Valor1", rows[0].Coluna1);
            Assert.Equal(new DateTime(2026, 1, 1), rows[0].Coluna2);
            Assert.Equal(DateTimeKind.Unspecified, rows[0].Coluna2.Kind);
            Assert.Equal(1, rows[0].Coluna3);
            Assert.Equal(0.1, rows[0].Coluna16, 10);
        }

        [Fact]
        public void PropertyNameIsTheDefaultColumnName()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            DefaultName[] rows = workbook.Sheet(0).Read<DefaultName>();
            Assert.Equal("Valor1", rows[0].Coluna1);
        }

        [Fact]
        public void AliasResolves()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Aliased[] rows = workbook.Sheet(0).Read<Aliased>();
            Assert.Equal("Valor1", rows[0].Value);
        }

        [Fact]
        public void IgnoredPropertyIsNotRead()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Ignored[] rows = workbook.Sheet(0).Read<Ignored>();
            Assert.Null(rows[0].Coluna1);
            Assert.Equal(1, rows[0].Coluna3);
        }

        [Fact]
        public void MissingOptionalColumnKeepsDefault()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            WithMissing[] rows = workbook.Sheet(0).Read<WithMissing>();
            Assert.Equal(RealDataRows, rows.Length);
            Assert.Equal("Valor1", rows[0].Coluna1);
            Assert.All(rows, r => Assert.Equal(0, r.Missing));
        }

        [Fact]
        public void MissingRequiredColumnThrows()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).Read<RequiredMissing>());
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
            Assert.Contains("NoSuchColumn", error.Message, StringComparison.Ordinal);
            Assert.Contains(nameof(RequiredMissing.Missing), error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void NoMatchingColumnThrows()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).Read<NoMatch>());
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
            Assert.Contains(nameof(NoMatch), error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void NullableGetsNullForEmptyCells()
        {
            using ExcelWorkbook workbook = OpenCsv(QtyCsv);
            NullableQty[] rows = workbook.Sheet(0).Read<NullableQty>();
            Assert.Equal(OneThenNull, rows.Select(r => r.Qty).ToArray());
        }

        [Fact]
        public void NonNullableGetsDefaultForEmptyCells()
        {
            using ExcelWorkbook workbook = OpenCsv(QtyCsv);
            PlainQty[] rows = workbook.Sheet(0).Read<PlainQty>();
            Assert.Equal(OneThenZero, rows.Select(r => r.Qty).ToArray());
        }

        [Fact]
        public void RequiredRejectsEmptyCell()
        {
            using ExcelWorkbook workbook = OpenCsv(QtyCsv);
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).Read<RequiredQty>());
            Assert.Equal(ExcelStatus.InvalidArgument, error.Status);
            Assert.Contains("Qty", error.Message, StringComparison.Ordinal);
            Assert.Contains("row 2", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void RequiredRejectsEmptyText()
        {
            using ExcelWorkbook workbook = OpenCsv("Name,Qty\r\na,1\r\n,2\r\n");
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).Read<RequiredName>());
            Assert.Contains("Name", error.Message, StringComparison.Ordinal);
            Assert.Contains("row 2", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void RequiredAllowEmptyAccepts()
        {
            using ExcelWorkbook workbook = OpenCsv(QtyCsv);
            RequiredAllowEmptyQty[] rows = workbook.Sheet(0).Read<RequiredAllowEmptyQty>();
            Assert.Equal(OneThenZero, rows.Select(r => r.Qty).ToArray());
        }

        [Fact]
        public void EmptyStringIsNull()
        {
            using ExcelWorkbook workbook = OpenCsv("Name,Qty\r\n,1\r\nb,2\r\n");
            NameOnly[] rows = workbook.Sheet(0).Read<NameOnly>();
            Assert.Equal(NullThenB, rows.Select(r => r.Name).ToArray());
        }

        [Fact]
        public void NonAsciiTextRoundTrips()
        {
            byte[] bom = new UTF8Encoding(true).GetPreamble();
            byte[] body = Encoding.UTF8.GetBytes("Name\r\nação 😀\r\n");
            using ExcelWorkbook workbook = ExcelWorkbook.Open(bom.Concat(body).ToArray(), new ExcelOpenOptions { Format = ExcelFormat.Csv });
            NameOnly[] rows = workbook.Sheet(0).Read<NameOnly>();
            Assert.Equal("ação 😀", Assert.Single(rows).Name);
        }

        [Fact]
        public void IntOverflowThrows()
        {
            using ExcelWorkbook workbook = OpenCsv("Big\r\n3000000000\r\n");
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).Read<BigInt>());
            Assert.Contains("Big", error.Message, StringComparison.Ordinal);
            Assert.Contains("row 1", error.Message, StringComparison.Ordinal);
            Assert.Equal(3000000000L, Assert.Single(workbook.Sheet(0).Read<BigLong>()).Big);
        }

        [Fact]
        public void EverySupportedTypeReads()
        {
            using ExcelWorkbook workbook = OpenCsv(
                "Flag,Ratio,Real,Small,Tiny,Money,At,Time,Count,Missing\r\n" +
                "true,1.5,2.25,-7,200,12.5,2026-01-02 03:04:05,12:30:15,5,\r\n");
            AllTypes row = Assert.Single(workbook.Sheet(0).Read<AllTypes>());
            Assert.True(row.Flag);
            Assert.Equal(1.5, row.Ratio);
            Assert.Equal(2.25f, row.Real);
            Assert.Equal((short)-7, row.Small);
            Assert.Equal((byte)200, row.Tiny);
            Assert.Equal(12.5m, row.Money);
            Assert.Equal(new DateTime(2026, 1, 2, 3, 4, 5), row.At);
            Assert.Equal(new TimeSpan(12, 30, 15), row.Time);
            Assert.Equal(5L, row.Count);
            Assert.Null(row.Missing);
        }

        [Fact]
        public void UnsupportedTypeThrowsBeforeReading()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            NotSupportedException error = Assert.Throws<NotSupportedException>(() => workbook.Sheet(0).Read<WithGuid>());
            Assert.Contains(nameof(WithGuid.Id), error.Message, StringComparison.Ordinal);
            Assert.Contains(nameof(Guid), error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void TableReadUsesTableRange()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", "tables.xlsx"));
            ExcelSheet table = workbook.Table("Sales");
            ProductOnly[] rows = table.Read<ProductOnly>();
            Assert.Equal(table.ReadRows().Count() - 1, rows.Length);
            string? first = table.ReadRows().ElementAt(1).Cells.Single(c => c.Column == 2).Text;
            Assert.NotNull(first);
            Assert.Equal(first, rows[0].Product);
        }

        [Fact]
        public void HeaderRowBelowOneIsRejected()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.Sheet(0).Read<Row>(headerRow: 0));
        }

        [Fact]
        public void NegativeParallelismIsRejected()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.Sheet(0).Read<Row>(degreeOfParallelism: -1));
        }

        [Fact]
        public void ParallelAndSerialAgree()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            Row[] serial = workbook.Sheet(0).Read<Row>(degreeOfParallelism: 1);
            Row[] parallel = workbook.Sheet(0).Read<Row>(degreeOfParallelism: 0);
            Assert.Equal(serial.Length, parallel.Length);
            for (int i = 0; i < serial.Length; i++)
            {
                Assert.Equal(serial[i].Coluna1, parallel[i].Coluna1);
                Assert.Equal(serial[i].Coluna2, parallel[i].Coluna2);
                Assert.Equal(serial[i].Coluna3, parallel[i].Coluna3);
                Assert.Equal(serial[i].Coluna16, parallel[i].Coluna16);
            }
        }

        private static ExcelWorkbook OpenCsv(string csv)
        {
            return ExcelWorkbook.Open(Encoding.UTF8.GetBytes(csv), new ExcelOpenOptions { Format = ExcelFormat.Csv });
        }

        public sealed class DefaultName
        {
            public string? Coluna1 { get; set; }
        }

        public sealed class Aliased
        {
            [ExcelColumn("ThisColumnDoesNotExist")]
            [ExcelColumn("Coluna1")]
            public string? Value { get; set; }
        }

        public sealed class Ignored
        {
            [ExcelIgnore]
            public string? Coluna1 { get; set; }

            public int Coluna3 { get; set; }
        }

        public sealed class WithMissing
        {
            public string? Coluna1 { get; set; }

            [ExcelColumn("NoSuchColumn")]
            public int Missing { get; set; }
        }

        public sealed class RequiredMissing
        {
            public string? Coluna1 { get; set; }

            [ExcelRequired]
            [ExcelColumn("NoSuchColumn")]
            public int Missing { get; set; }
        }

        public sealed class NoMatch
        {
            [ExcelColumn("NoSuchColumn")]
            public string? Value { get; set; }
        }

        public sealed class NullableQty
        {
            public int? Qty { get; set; }
        }

        public sealed class PlainQty
        {
            public int Qty { get; set; }
        }

        public sealed class RequiredQty
        {
            [ExcelRequired]
            public int Qty { get; set; }
        }

        public sealed class RequiredAllowEmptyQty
        {
            [ExcelRequired(AllowEmpty = true)]
            public int Qty { get; set; }
        }

        public sealed class NameOnly
        {
            public string? Name { get; set; }
        }

        public sealed class RequiredName
        {
            [ExcelRequired]
            public string? Name { get; set; }
        }

        public sealed class BigInt
        {
            public int Big { get; set; }
        }

        public sealed class BigLong
        {
            public long Big { get; set; }
        }

        public sealed class AllTypes
        {
            public bool Flag { get; set; }

            public double Ratio { get; set; }

            public float Real { get; set; }

            public short Small { get; set; }

            public byte Tiny { get; set; }

            public decimal Money { get; set; }

            public DateTime At { get; set; }

            public TimeSpan Time { get; set; }

            public long? Count { get; set; }

            public DateTime? Missing { get; set; }
        }

        public sealed class WithGuid
        {
            public Guid Id { get; set; }

            public string? Coluna1 { get; set; }
        }

        public sealed class ProductOnly
        {
            public string? Product { get; set; }
        }
    }
}
