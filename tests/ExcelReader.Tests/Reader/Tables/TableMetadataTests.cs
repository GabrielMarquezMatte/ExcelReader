using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Tables
{
    public class TableMetadataTests
    {
        private const string Rows = """<row r="1"><c r="A1" t="inlineStr"><is><t>H</t></is></c></row><row r="2"><c r="A2"><v>1</v></c></row>""";

        public static TheoryData<string> Fixtures { get; } = new() { "tables.xlsx", "tables.xlsb" };

        private static readonly (string Name, int Sheet, string Ref, int FirstRow, int FirstColumn, int LastRow, int LastColumn,
            int HeaderRowCount, int TotalsRowCount, string[] Columns, string? Style)[] Expected =
        [
            ("Sales", 0, "C4:E8", 3, 2, 7, 4, 1, 1, ["Product", "Qty", "Price"], "TableStyleMedium2"),
            ("Stock", 0, "G4:H7", 3, 6, 6, 7, 1, 0, ["Product", "Qty"], "TableStyleLight9"),
            ("NoHeader", 1, "A2:B4", 1, 0, 3, 1, 0, 0, ["Letter", "Number"], "TableStyleMedium2"),
            ("NoStyle", 1, "D1:E3", 0, 3, 2, 4, 1, 0, ["Name", "Value"], null),
        ];

        internal static string FixturePath(string file)
        {
            return Path.Combine(AppContext.BaseDirectory, "data", "tables", file);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void FixtureTablesReportTheirMetadata(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(FixturePath(file));

            AssertExpectedTables(workbook);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public async Task AsyncOpenReportsTheSameTables(string file)
        {
            await using IExcelWorkbook workbook = await Excel.OpenAsync(
                File.OpenRead(FixturePath(file)), leaveOpen: false, ct: TestContext.Current.CancellationToken);

            AssertExpectedTables(workbook);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void TryGetTableIgnoresCase(string file)
        {
            using IExcelWorkbook workbook = Excel.Open(FixturePath(file));

            Assert.True(workbook.TryGetTable("sales", out ExcelTable? sales));
            Assert.Equal("Sales", sales.Name);
            Assert.Equal("Data", sales.Sheet.Name);
            Assert.False(workbook.TryGetTable("Missing", out _));
        }

        [Fact]
        public void EncryptedWorkbookReportsItsTables()
        {
            using MemoryStream encrypted = new();
            using (FileStream package = File.OpenRead(FixturePath("tables.xlsx")))
            {
                Excel.EncryptPackage(package, encrypted, "hunter2");
            }
            encrypted.Position = 0;

            using IExcelWorkbook workbook = Excel.Open(encrypted, options: ExcelReaderOptions.Default with { Password = "hunter2" });

            AssertExpectedTables(workbook);
        }

        [Fact]
        public void PrefixedTablePartParses()
        {
            string table = """<x:table xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="T" displayName="T" ref="A1:A2"><x:tableColumns count="1"><x:tableColumn id="1" name="H"/></x:tableColumns><x:tableStyleInfo name="TableStyleLight1"/></x:table>""";
            using MemoryStream ms = TableWorkbooks.Xlsx(Rows, table);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            ExcelTable parsed = Assert.Single(workbook.Tables);
            Assert.Equal(("T", "A1:A2", "TableStyleLight1"), (parsed.Name, parsed.Ref, parsed.StyleName));
            Assert.Equal(["H"], parsed.ColumnNames);
        }

        [Fact]
        public void EscapedColumnNamesAreDecoded()
        {
            using MemoryStream ms = TableWorkbooks.Xlsx(Rows, TableWorkbooks.Table("A1:A2", """<tableColumn id="1" name="Qty &amp; Price"/>"""));
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Equal(["Qty & Price"], Assert.Single(workbook.Tables).ColumnNames);
        }

        [Fact]
        public void OnlyTableRelationshipsAreFollowed()
        {
            string others =
                """<Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>"""
                + """<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>""";
            using MemoryStream ms = TableWorkbooks.Xlsx(Rows, TableWorkbooks.Table("A1:A2", """<tableColumn id="1" name="H"/>"""), extraRelationships: others);
            using XlsxWorkbook workbook = Excel.FromXlsx(ms);

            Assert.Equal("T", Assert.Single(workbook.Tables).Name);
        }

        private static void AssertExpectedTables(IExcelWorkbook workbook)
        {
            Assert.Equal(Expected.Length, workbook.Tables.Count);
            for (int i = 0; i < Expected.Length; i++)
            {
                ExcelTable table = workbook.Tables[i];
                var expected = Expected[i];
                Assert.Equal((expected.Name, expected.Sheet, expected.Ref), (table.Name, table.Sheet.Index, table.Ref));
                Assert.Equal((expected.FirstRow, expected.FirstColumn, expected.LastRow, expected.LastColumn),
                    (table.FirstRow, table.FirstColumn, table.LastRow, table.LastColumn));
                Assert.Equal((expected.HeaderRowCount, expected.TotalsRowCount), (table.HeaderRowCount, table.TotalsRowCount));
                Assert.Equal(expected.Columns, table.ColumnNames);
                Assert.Equal(expected.Style, table.StyleName);
            }
        }
    }
}
