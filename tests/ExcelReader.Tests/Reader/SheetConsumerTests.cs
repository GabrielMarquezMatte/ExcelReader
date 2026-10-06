using System.Data;
using System.Data.Common;
using System.Globalization;
using Apache.Arrow;
using ExcelReader.Arrow;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using ExcelReader.Core.Reader.Schema;
using ExcelReader.Core.Reader.Xls;
using ExcelReader.Core.Reader.Xlsb;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer;
using ExcelReader.Core.Writer.Xls;
using ExcelReader.Core.Writer.Xlsb;
using ExcelReader.Core.Writer.Xlsx;

namespace ExcelReader.Tests.Reader
{
    public sealed class SheetConsumerTests
    {
        public sealed class Item
        {
            [ExcelColumn("Name")]
            public string Name { get; set; } = "";

            [ExcelColumn("Amount")]
            public int Amount { get; set; }
        }

        private const int Rows = 20;

        private static void Fill<TSheet, TRow>(IWorkbookWriter<TSheet> workbook)
            where TSheet : ISheetWriter<TRow>
            where TRow : IRowWriter
        {
            for (int s = 0; s < 2; s++)
            {
                TSheet sheet = workbook.AddSheet("S" + s.ToString(CultureInfo.InvariantCulture));
                using (TRow header = sheet.StartRow())
                {
                    header.Write("Name");
                    header.Write("Amount");
                }
                for (int r = 0; r < Rows; r++)
                {
                    using TRow row = sheet.StartRow();
                    row.Write("s" + s.ToString(CultureInfo.InvariantCulture) + "-" + r.ToString(CultureInfo.InvariantCulture));
                    row.Write((s * 100) + r);
                }
                sheet.End();
                sheet.Dispose();
            }
            workbook.End();
        }

        private static byte[] BuildXlsx()
        {
            using MemoryStream buffer = new();
            using (XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                Fill<XlsxSheetWriter, XlsxRowWriter>(workbook);
            }
            return buffer.ToArray();
        }

        private static byte[] BuildXlsb()
        {
            using MemoryStream buffer = new();
            using (XlsbWorkbookWriter workbook = XlsbWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                Fill<XlsbSheetWriter, XlsbRowWriter>(workbook);
            }
            return buffer.ToArray();
        }

        private static byte[] BuildXls()
        {
            using MemoryStream buffer = new();
            using (XlsWorkbookWriter workbook = XlsWorkbookWriter.Create(buffer, leaveOpen: true))
            {
                Fill<XlsSheetWriter, XlsRowWriter>(workbook);
            }
            return buffer.ToArray();
        }

        private static byte[] BuildCsv()
        {
            System.Text.StringBuilder text = new("Name,Amount\n");
            for (int r = 0; r < Rows; r++)
            {
                text.Append("s1-").Append(r.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append((100 + r).ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return System.Text.Encoding.UTF8.GetBytes(text.ToString());
        }

        private static void AssertSecondSheet(IEnumerable<Item> items)
        {
            List<Item> list = [.. items];
            Assert.Equal(Rows, list.Count);
            for (int r = 0; r < Rows; r++)
            {
                Assert.Equal("s1-" + r.ToString(CultureInfo.InvariantCulture), list[r].Name);
                Assert.Equal(100 + r, list[r].Amount);
            }
        }

        [Fact]
        public void Parse_Reads_The_Sheet_It_Is_Given_For_Every_Typed_Overload()
        {
            ExcelParser<Item> parser = ExcelParser.FromAttributes<Item>();

            using XlsxWorkbook xlsx = Excel.FromXlsx(BuildXlsx());
            AssertSecondSheet(parser.Parse(xlsx.Sheets[1]));

            using XlsbWorkbook xlsb = Excel.FromXlsb(BuildXlsb());
            AssertSecondSheet(parser.Parse(xlsb.Sheets[1]));

            using XlsWorkbook xls = Excel.FromXls(BuildXls());
            AssertSecondSheet(parser.Parse(xls.Sheets[1]));

            using CsvReader csv = Excel.FromCsv(BuildCsv());
            AssertSecondSheet(parser.Parse(csv.Sheets[0]));
        }

        [Fact]
        public async Task Parse_Over_A_Format_Agnostic_Sheet_Works_Sync_And_Async()
        {
            ExcelParser<Item> parser = ExcelParser.FromAttributes<Item>();
            using IExcelWorkbook workbook = Excel.Open(BuildXlsx());
            IExcelSheet sheet = workbook.SheetAt(1);

            AssertSecondSheet(parser.Parse(sheet));

            List<Item> asyncItems = [];
            await foreach (Item item in parser.Parse(sheet).WithCancellation(TestContext.Current.CancellationToken))
            {
                asyncItems.Add(item);
            }
            AssertSecondSheet(asyncItems);
        }

        [Fact]
        public void One_Parse_Result_Can_Be_Enumerated_Twice_And_Sheets_Stay_Independent()
        {
            ExcelParser<Item> parser = ExcelParser.FromAttributes<Item>();
            using XlsxWorkbook workbook = Excel.FromXlsx(BuildXlsx());
            var second = parser.Parse(workbook.Sheets[1]);
            var first = parser.Parse(workbook.Sheets[0]);

            AssertSecondSheet(second);
            Assert.Equal("s0-0", first.First().Name);
            AssertSecondSheet(second);
        }

        [Fact]
        public void InferSchema_Samples_The_Sheet_It_Is_Given()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(BuildXlsx());
            ExcelColumnSchema[] schema = Excel.InferSchema(workbook.Sheets[1]);
            Assert.Equal(2, schema.Length);
            Assert.Equal("Name", schema[0].Name);
            Assert.Equal(ExcelColumnType.Int64Column, schema[1].Type);
            Assert.Equal(schema.Length, Excel.InferSchema(workbook.Sheets[1], 1, 5).Length);
            Assert.Throws<ArgumentNullException>(() => Excel.InferSchema((IExcelSheet)null!));
        }

        [Fact]
        public void ExcelDataReader_Reads_The_Sheet_It_Is_Given_And_Names_Its_Table_After_It()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(BuildXlsx());
            using ExcelDataReader data = new(workbook.Sheets[1]);
            Assert.Equal(2, data.FieldCount);
            Assert.True(data.Read());
            Assert.Equal("s1-0", data.GetString(0));
            using DataTable schema = data.GetSchemaTable()!;
            Assert.Equal("S1", schema.Rows[0][SchemaTableColumn.BaseTableName]);
            Assert.Throws<ArgumentNullException>(() => new ExcelDataReader((IExcelSheet)null!));
        }

        [Fact]
        public void ToArrowRecordBatch_Converts_The_Sheet_It_Is_Given()
        {
            using XlsxWorkbook workbook = Excel.FromXlsx(BuildXlsx());
            using RecordBatch batch = workbook.Sheets[1].ToArrowRecordBatch();
            Assert.Equal(Rows, batch.Length);
            Assert.Equal(2, batch.ColumnCount);
            Assert.Equal("s1-0", ((StringArray)batch.Column(0)).GetString(0));
        }
    }
}
