using System.Globalization;
using System.Text;
using ExcelReader.NetFramework.Reader;

namespace ExcelReader.NetFramework.Tests.Reader
{
    public sealed class ForEachAndAggregateTests
    {
        private const int RealRowCount = 101;
        private const int CsvRowCount = 200_000;
        private static readonly string?[] CsvTexts = { "Name", "Qty", "a", "1.5" };

        [Theory]
        [InlineData("RealExcel.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void ForEachSeesTheRowsReadRowsSees(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path(file));
            ExcelSheet sheet = workbook.Sheet(0);
            List<ExcelRow> seen = new List<ExcelRow>();
            sheet.ForEach(row => seen.Add(row.ToRow()));

            List<ExcelRow> expected = sheet.ReadRows().ToList();
            Assert.Equal(RealRowCount, seen.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i].Cells.Select(c => c.Column), seen[i].Cells.Select(c => c.Column));
                Assert.Equal(expected[i].Cells.Select(c => c.Type), seen[i].Cells.Select(c => c.Type));
            }

            Assert.Contains(seen[0].Cells, c => string.Equals(c.Text, "Coluna1", StringComparison.Ordinal));
        }

        [Fact]
        public void ForEachOverATableStartsAtItsHeader()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("tables", "tables.xlsx"));
            List<string?> firstCells = new List<string?>();
            workbook.Table("Sales").ForEach(row => firstCells.Add(row[0].GetText()));
            Assert.Equal(4, firstCells.Count);
            Assert.Equal("Product", firstCells[0]);
        }

        [Fact]
        public void ForEachRethrowsAndStops()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            int calls = 0;
            InvalidTimeZoneException error = Assert.Throws<InvalidTimeZoneException>(() => workbook.Sheet(0).ForEach(_ =>
            {
                calls++;
                if (calls == 3)
                {
                    throw new InvalidTimeZoneException("stop");
                }
            }));
            Assert.Equal("stop", error.Message);
            Assert.Equal(3, calls);
        }

        [Fact]
        public void CellViewExposesNumberAndText()
        {
            using ExcelWorkbook workbook = OpenCsv("Name,Qty\r\na,1.5\r\n");
            List<double?> numbers = new List<double?>();
            List<string?> texts = new List<string?>();
            workbook.Sheet(0).ForEach(row =>
            {
                for (int i = 0; i < row.Count; i++)
                {
                    numbers.Add(row[i].Number);
                    texts.Add(row[i].GetText());
                }
            });

            Assert.Equal(new double?[] { null, null, null, 1.5 }, numbers);
            Assert.Equal(CsvTexts, texts, StringComparer.Ordinal);
        }

        [Fact]
        public void RowViewIndexerRejectsOutOfRange()
        {
            using ExcelWorkbook workbook = OpenCsv("a,b\r\n");
            int count = -1;
            Assert.Throws<ArgumentOutOfRangeException>(() => workbook.Sheet(0).ForEach(row =>
            {
                count = row.Count;
                _ = row[row.Count];
            }));
            Assert.Equal(2, count);
        }

        [Theory]
        [InlineData("RealExcel.xlsx")]
        [InlineData("RealExcel.xlsb")]
        public void AggregateCountsDataRows(string file)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path(file));
            ExcelSheet sheet = workbook.Sheet(0);
            Assert.Equal(RealRowCount - 1, sheet.Aggregate(() => new RowCount()).Rows);
            Assert.Equal(RealRowCount, sheet.Aggregate(() => new RowCount(), headerRow: 0).Rows);
        }

        [Theory]
        [InlineData(1)]
        [InlineData(4)]
        [InlineData(0)]
        public void AggregateCombinesChunksInSourceOrder(int degreeOfParallelism)
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(LargeCsv(), new ExcelOpenOptions { Format = ExcelFormat.Csv });
            Ordered result = workbook.Sheet(0).Aggregate(() => new Ordered(), degreeOfParallelism: degreeOfParallelism);

            Assert.Equal(CsvRowCount, result.Ids.Count);
            for (int i = 0; i < result.Ids.Count; i++)
            {
                Assert.Equal(i, result.Ids[i]);
            }

            Assert.Equal((double)CsvRowCount * (CsvRowCount - 1) / 2, result.Sum);
        }

        [Fact]
        public void AggregateRethrowsTheCallbackException()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(LargeCsv(), new ExcelOpenOptions { Format = ExcelFormat.Csv });
            InvalidTimeZoneException error = Assert.Throws<InvalidTimeZoneException>(() => workbook.Sheet(0).Aggregate(() => new Throwing()));
            Assert.Equal("boom", error.Message);
        }

        [Fact]
        public void AggregateRejectsANullSeedResult()
        {
            using ExcelWorkbook workbook = OpenCsv("a\r\n1\r\n");
            Assert.Throws<InvalidOperationException>(() => workbook.Sheet(0).Aggregate<RowCount>(() => null!));
        }

        [Fact]
        public void AggregateValidatesArguments()
        {
            using ExcelWorkbook workbook = OpenCsv("a\r\n1\r\n");
            ExcelSheet sheet = workbook.Sheet(0);
            Assert.Throws<ArgumentNullException>(() => sheet.Aggregate<RowCount>(null!));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.Aggregate(() => new RowCount(), headerRow: -1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.Aggregate(() => new RowCount(), degreeOfParallelism: -1));
            Assert.Throws<ArgumentNullException>(() => sheet.ForEach(null!));
        }

        private static ExcelWorkbook OpenCsv(string csv)
        {
            return ExcelWorkbook.Open(Encoding.UTF8.GetBytes(csv), new ExcelOpenOptions { Format = ExcelFormat.Csv });
        }

        private static byte[] LargeCsv()
        {
            StringBuilder csv = new StringBuilder("Id,Value\r\n");
            for (int i = 0; i < CsvRowCount; i++)
            {
                csv.Append(i.ToString(CultureInfo.InvariantCulture)).Append(",x").Append(i.ToString(CultureInfo.InvariantCulture)).Append("\r\n");
            }

            return Encoding.UTF8.GetBytes(csv.ToString());
        }

        private sealed class RowCount : IExcelAccumulator<RowCount>
        {
            internal long Rows { get; private set; }

            public void Accumulate(ExcelRowView row)
            {
                Rows++;
            }

            public void Combine(RowCount other)
            {
                Rows += other.Rows;
            }
        }

        private sealed class Ordered : IExcelAccumulator<Ordered>
        {
            internal List<int> Ids { get; } = new List<int>();

            internal double Sum { get; private set; }

            public void Accumulate(ExcelRowView row)
            {
                double id = row[0].Number!.Value;
                Ids.Add((int)id);
                Sum += id;
            }

            public void Combine(Ordered other)
            {
                Ids.AddRange(other.Ids);
                Sum += other.Sum;
            }
        }

        private sealed class Throwing : IExcelAccumulator<Throwing>
        {
            public void Accumulate(ExcelRowView row)
            {
                throw new InvalidTimeZoneException("boom");
            }

            public void Combine(Throwing other)
            {
            }
        }
    }
}
