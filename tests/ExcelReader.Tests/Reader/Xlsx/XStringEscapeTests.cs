using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;
using ExcelReader.Core.Writer.Xlsx;
using ExcelReader.Tests.Writer.Tables;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class XStringEscapeTests
    {
        public static TheoryData<string> RoundTripText()
        {
            return new()
            {
                "a\rb",
                "line1\r\nline2",
                "a\u0001b _x0041_",
                "_x000D_",
                "x_y_z",
                "trailing_",
                "_x0041",
                "\u001F\u0000",
            };
        }

        private static List<string> ReadFirstColumn(Stream workbook)
        {
            using XlsxWorkbook reader = Excel.FromXlsx(workbook);
            var values = new List<string>();
            foreach (Row row in reader.FirstSheet)
            {
                values.Add(row[0].GetString());
            }
            return values;
        }

        [Theory]
        [MemberData(nameof(RoundTripText))]
        public async Task TextWrittenByTheLibraryReadsBackUnchanged(string text)
        {
            foreach (bool shared in new[] { false, true })
            {
                var ms = new MemoryStream();
                await using (XlsxWorkbookWriter wb = XlsxWorkbookWriter.Create(ms, leaveOpen: true, new XlsxWriterOptions { UseSharedStrings = shared }))
                {
                    XlsxSheetWriter sheet = wb.AddSheet("S1");
                    await using (XlsxRowWriter row = await sheet.StartRowAsync(TestContext.Current.CancellationToken))
                    {
                        row.Write(text);
                    }
                    await sheet.EndAsync(TestContext.Current.CancellationToken);
                    await wb.EndAsync(TestContext.Current.CancellationToken);
                }
                ms.Position = 0;

                Assert.Equal([text], ReadFirstColumn(ms));
            }
        }

        public static TheoryData<string, string> ExcelEscapes()
        {
            return new()
            {
                { "a_x000D_b", "a\rb" },
                { "a_x000d_b", "a\rb" },
                { "_x000D__x000A_", "\r\n" },
                { "_x005F_x0041_", "_x0041_" },
                { "_x0041_", "A" },
                { "_xD83D__xDE00_", "\U0001F600" },
                { "_xD800_", "_xD800_" },
                { "_xDE00_x", "_xDE00_x" },
                { "_X000D_", "_X000D_" },
                { "_x00D_", "_x00D_" },
                { "_x000G_", "_x000G_" },
            { "_x00!0_", "_x00!0_" },
            { "_x00@0_", "_x00@0_" },
                { "a_x000D", "a_x000D" },
                { "&amp;_x0026_", "&&" },
            };
        }

        [Theory]
        [MemberData(nameof(ExcelEscapes))]
        public void EscapesInSharedInlineAndFormulaTextAreDecoded(string xml, string expected)
        {
            string rows =
                $"""<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>{xml}</t></is></c><c r="C1" t="str"><f>A1</f><v>{xml}</v></c><c r="D1" t="s"><v>1</v></c></row>""";
            string sharedStrings = $"""<si><t>{xml}</t></si><si><r><t>[</t></r><r><t>{xml}</t></r><r><t>]</t></r></si>""";
            using MemoryStream ms = WorkbookBuilder.Build(rows, sharedStrings);
            using XlsxWorkbook reader = Excel.FromXlsx(ms);

            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();
            Assert.True(e.MoveNext());
            Row row = e.Current;
            Assert.Equal(expected, row[0].GetString());
            Assert.Equal(expected, row[1].GetString());
            Assert.Equal(expected, row[2].GetString());
            Assert.Equal("[" + expected + "]", row[3].GetString());
        }

        [Fact]
        public void APrefixedSheetDecodesEscapesInInlineText()
        {
            using MemoryStream ms = WorkbookBuilder.BuildPrefixed(
                "x", """<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>a_x000D_b</x:t></x:is></x:c></x:row>""");

            Assert.Equal(["a\rb"], ReadFirstColumn(ms));
        }

        [Fact]
        public void ATableColumnNameWrittenByTheLibraryReadsBackUnchanged()
        {
            ITableBook book = TableBooks.Create("xlsx");
            book.AddSheet("S");
            book.BeginTable("Codes", ["Code_x0041_", "Line_x000D_"]);
            book.Row("a", "b");
            byte[] package = book.Finish();

            using XlsxWorkbook reader = Excel.FromXlsx(package);
            Assert.True(reader.TryGetTable("Codes", out ExcelTable? table));
            Assert.Equal(["Code_x0041_", "Line_x000D_"], table.ColumnNames);
        }
    }
}
