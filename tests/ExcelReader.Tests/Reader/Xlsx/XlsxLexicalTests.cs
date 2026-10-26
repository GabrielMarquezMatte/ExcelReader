using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public sealed class XlsxLexicalTests
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static string DateStyles(string xf, string numFmts = "")
        {
            return $"""<styleSheet>{numFmts}<cellXfs count="2"><xf numFmtId="0"/>{xf}</cellXfs></styleSheet>""";
        }

        private static byte[] Build(string rows, string? styles = null)
        {
            using MemoryStream built = WorkbookBuilder.Build(rows, styles: styles);
            return built.ToArray();
        }

        private static string Line(Row row)
        {
            List<string> cells = [];
            for (int column = 0; column < row.ColumnCount; column++)
            {
                Cell cell = row[column];
                cells.Add($"{cell.Type}:{cell.GetString()}");
            }
            return string.Join("|", cells);
        }

        // Reads the first row from memory, a non-seekable stream, and asynchronously, which must all agree.
        private static async Task<string> FirstRow(byte[] bytes)
        {
            string fromMemory;
            using (IExcelWorkbook memory = Excel.FromXlsx(bytes))
            using (IExcelRowEnumerator rows = memory.FirstSheet.GetEnumerator())
            {
                Assert.True(rows.MoveNext());
                fromMemory = Line(rows.Current);
            }

            string fromStream;
            using (IExcelWorkbook stream = Excel.FromXlsx(new NonSeekableStream(bytes), leaveOpen: false))
            using (IExcelRowEnumerator rows = stream.FirstSheet.GetEnumerator())
            {
                Assert.True(rows.MoveNext());
                fromStream = Line(rows.Current);
            }

            using IExcelWorkbook asyncReader = Excel.FromXlsx(new MemoryStream(bytes, writable: false), leaveOpen: false);
            await using IExcelRowEnumerator asyncRows = asyncReader.FirstSheet.GetAsyncEnumerator(TestContext.Current.CancellationToken);
            Assert.True(await asyncRows.MoveNextAsync());
            Assert.Equal(fromMemory, fromStream);
            Assert.Equal(fromMemory, Line(asyncRows.Current));
            return fromMemory;
        }

        [Theory]
        [InlineData(" ", " ")]
        [InlineData("\t", "\t")]
        [InlineData("\r\n", "\n")]
        [InlineData("", "  ")]
        public async Task Should_KeepCellCoordinates_When_AttributesHaveWhitespaceAroundEquals(string before, string after)
        {
            byte[] bytes = Build($"""<row r{before}={after}"1"><c r="A1" t="inlineStr"><is><t>A</t></is></c><c r{before}={after}"C1"{"\n"}t{before}={after}"inlineStr"><is><t>C</t></is></c></row>""");

            Assert.Equal("ExcelString:A|Empty:|ExcelString:C", await FirstRow(bytes));
        }

        [Theory]
        [InlineData("<xf numFmtId=\"14\"/>")]
        [InlineData("<xf\nnumFmtId=\"14\"/>")]
        [InlineData("<xf\tnumFmtId = \"14\"/>")]
        [InlineData("<xf\r\n numFmtId=\r\n'14'/>")]
        public async Task Should_DetectDateStyles_When_TheXfTagUsesAnyWhitespace(string xf)
        {
            byte[] bytes = Build("""<row r="1"><c r="A1" s="1"><v>45000</v></c></row>""", DateStyles(xf));

            Assert.StartsWith("Date:", await FirstRow(bytes), StringComparison.Ordinal);
        }

        [Fact]
        public async Task Should_DetectADateFormat_When_ItsFormatCodeHoldsAGreaterThanSign()
        {
            string numFmts = """<numFmts count="1"><numFmt numFmtId="164" formatCode="[>1]yyyy-mm-dd"/></numFmts>""";
            byte[] bytes = Build("""<row r="1"><c r="A1" s="1"><v>45000</v></c></row>""", DateStyles("""<xf numFmtId="164"/>""", numFmts));

            Assert.StartsWith("Date:", await FirstRow(bytes), StringComparison.Ordinal);
        }

        [Fact]
        public async Task Should_IgnoreACommentedOutNumberFormat()
        {
            string numFmts = """<numFmts count="1"><!-- <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/> --><numFmt numFmtId="164" formatCode="0.00"/></numFmts>""";
            byte[] bytes = Build("""<row r="1"><c r="A1" s="1"><v>45000</v></c></row>""", DateStyles("""<xf numFmtId="164"/>""", numFmts));

            Assert.StartsWith("Number:", await FirstRow(bytes), StringComparison.Ordinal);
        }

        [Fact]
        public void Should_IgnoreCommentedOutSheetsAndRelationships()
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] bytes = Rewrite(built.ToArray(), "xl/workbook.xml",
                $"""<workbook xmlns="{Main}" xmlns:r="{Rel}"><sheets><!-- <sheet name="Phantom" sheetId="9" r:id="missing"/> --><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""");

            using XlsxWorkbook reader = Excel.FromXlsx(bytes);
            Assert.Equal(1, reader.SheetCount);
            Assert.Equal("S1", reader.FirstSheet.Name);
        }

        [Theory]
        [InlineData("./worksheets/sheet1.xml")]
        [InlineData("../xl/worksheets/sheet1.xml")]
        public async Task Should_ReadTheSheet_When_ItsRelationshipTargetHasDotSegments(string target)
        {
            using MemoryStream built = WorkbookBuilder.Build("""<row r="1"><c r="A1"><v>1</v></c></row>""");
            byte[] bytes = Rewrite(built.ToArray(), "xl/_rels/workbook.xml.rels",
                $"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="x" Target="{target}"/></Relationships>""");

            Assert.Equal("Number:1", await FirstRow(bytes));
        }

        private static byte[] Rewrite(byte[] zip, string entry, string content)
        {
            using MemoryStream buffer = new();
            buffer.Write(zip);
            using (System.IO.Compression.ZipArchive archive = new(buffer, System.IO.Compression.ZipArchiveMode.Update, leaveOpen: true))
            {
                archive.GetEntry(entry)!.Delete();
                using StreamWriter writer = new(archive.CreateEntry(entry).Open());
                writer.Write(content);
            }
            return buffer.ToArray();
        }
    }
}
