using System.Text;

namespace ExcelReader.NetFramework.Tests
{
    public sealed class BatchAndSchemaTests
    {
        private static readonly int[] ExpectedBatchSizes = { 30, 30, 30, 10 };

        [Fact]
        public void BatchesCoverAllRows()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            List<List<Row>> batches = sheet.ReadBatches<Row>(30).ToList();
            Assert.Equal(ExpectedBatchSizes, batches.Select(b => b.Count).ToArray());

            List<Row> all = batches.SelectMany(b => b).ToList();
            List<Row> expected = sheet.Read<Row>();
            Assert.Equal(expected.Count, all.Count);
            for (int i = 0; i < expected.Count; i++)
            {
                Assert.Equal(expected[i].Coluna1, all[i].Coluna1);
                Assert.Equal(expected[i].Coluna2, all[i].Coluna2);
                Assert.Equal(expected[i].Coluna3, all[i].Coluna3);
                Assert.Equal(expected[i].Coluna16, all[i].Coluna16);
            }
        }

        [Fact]
        public void BatchLargerThanSheetIsOneBatch()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            List<List<Row>> batches = workbook.Sheet(0).ReadBatches<Row>(1000).ToList();
            Assert.Single(batches);
            Assert.Equal(100, batches[0].Count);
        }

        [Fact]
        public void BreakingOutOfBatchesReleasesReader()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            for (int i = 0; i < 200; i++)
            {
                foreach (List<Row> batch in sheet.ReadBatches<Row>(10))
                {
                    Assert.Equal(10, batch.Count);
                    break;
                }
            }

            Assert.Equal(100, sheet.Read<Row>().Count);
        }

        [Fact]
        public void BatchesSurviveWorkbookDispose()
        {
            ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            int total = 0;
            bool first = true;
            foreach (List<Row> batch in workbook.Sheet(0).ReadBatches<Row>(30))
            {
                total += batch.Count;
                if (first)
                {
                    workbook.Dispose();
                    first = false;
                }
            }

            Assert.Equal(100, total);
        }

        [Fact]
        public void RequiredViolationReportsAbsoluteDataRow()
        {
            byte[] csv = Encoding.UTF8.GetBytes("Name,Qty\r\na,1\r\nb,2\r\nc,3\r\n,4\r\ne,5\r\n");
            using ExcelWorkbook workbook = ExcelWorkbook.Open(csv, new ExcelOpenOptions { Format = ExcelFormat.Csv });
            ExcelReaderException error = Assert.Throws<ExcelReaderException>(() => workbook.Sheet(0).ReadBatches<RequiredName>(2).ToList());
            Assert.Contains("'Name' is empty in data row 4.", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void BatchRowsBelowOneIsRejected()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.ReadBatches<Row>(0));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.ReadBatches<Row>(10, 0));
        }

        [Fact]
        public void InfersSchemaFromHeader()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            IReadOnlyList<ExcelColumnSchema> schema = workbook.Sheet(0).InferSchema(1, 100);
            Assert.Equal("Coluna1", schema[0].Name);
            Assert.Equal(ExcelColumnType.String, schema[0].Type);
            ExcelColumnSchema third = schema.Single(c => string.Equals(c.Name, "Coluna3", StringComparison.Ordinal));
            Assert.Equal(ExcelColumnType.Int64, third.Type);
        }

        [Fact]
        public void InferSchemaLeavesSheetReadable()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            sheet.InferSchema();
            Assert.Equal(100, sheet.Read<Row>().Count);
        }

        [Fact]
        public void InferSchemaRejectsBadArguments()
        {
            using ExcelWorkbook workbook = ExcelWorkbook.Open(Fixtures.Path("RealExcel.xlsx"));
            ExcelSheet sheet = workbook.Sheet(0);
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.InferSchema(-1));
            Assert.Throws<ArgumentOutOfRangeException>(() => sheet.InferSchema(1, 0));
        }

        public sealed class RequiredName
        {
            [ExcelRequired]
            public string? Name { get; set; }
        }
    }
}
