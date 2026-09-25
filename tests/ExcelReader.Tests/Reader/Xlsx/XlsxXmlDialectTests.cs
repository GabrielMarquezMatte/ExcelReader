using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class XlsxXmlDialectTests
    {
        [Fact]
        public void SingleQuotedCellAttributesAreRead()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r='C1' t='s'><v>0</v></c></row>""",
                sharedStrings: "<si><t>single quoted</t></si>");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("single quoted", e.Current[2].GetString());
        }

        [Theory]
        [InlineData("""<c r="AB1" s="7" t="s"><v>0</v></c>""")]
        [InlineData("""<c t="s" s="7" r="AB1"><v>0</v></c>""")]
        [InlineData("""<c r="AB1"  s="7" t="s"><v>0</v></c>""")]
        [InlineData("""<c r="AB1" s='7' t="s"><v>0</v></c>""")]
        [InlineData("""<c r="AB1" s="7" t="s" vm="1"><v>0</v></c>""")]
        [InlineData("""<c r="AB1" s="7" t="s" ><v>0</v></c>""")]
        [InlineData("""<c r="AA1"/><c s="7" t="s"><v>0</v></c>""")]
        public void CellTagAttributeLayoutsResolveTheSameCell(string cells)
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1">{cells}</row>""",
                sharedStrings: "<si><t>hit</t></si>");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Cell cell = e.Current[27];
            Assert.Equal("hit", cell.GetString());
            Assert.Equal(7, cell.StyleIndex);
        }

        [Theory]
        [InlineData("1", "second")]
        [InlineData("0001", "second")]
        [InlineData("+1", "second")]
        [InlineData("0", "first")]
        [InlineData("2", "")]
        [InlineData("4294967297", "")]
        public void SharedStringIndexFormsResolveLikeTheGenericParser(string index, string expected)
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                $"""<row r="1"><c r="A1" t="s"><v>{index}</v></c><c r="B1"><v>7</v></c></row>""",
                sharedStrings: "<si><t>first</t></si><si><t>second</t></si>");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(expected, e.Current[0].GetString());
            Assert.Equal("7", e.Current[1].GetString());
        }

        [Fact]
        public void PlainAndMarkedUpSharedStringsInterleaveCorrectly()
        {
            string[] expected = ["plain", "a & b", " padded ", "rich text", "", "", "base", "cdata &amp;", "last"];
            using MemoryStream ms = WorkbookBuilder.Build(
                "<row r=\"1\">" + string.Concat(expected.Select((_, i) => $"<c t=\"s\"><v>{i}</v></c>")) + "</row>",
                sharedStrings: "<si><t>plain</t></si>"
                    + "<si><t>a &amp; b</t></si>"
                    + "<si><t xml:space=\"preserve\"> padded </t></si>"
                    + "<si><r><t>rich</t></r><r><t> text</t></r></si>"
                    + "<si><t></t></si>"
                    + "<si><t/></si>"
                    + "<si><t>base</t><rPh sb=\"0\" eb=\"1\"><t>ruby</t></rPh></si>"
                    + "<si><t><![CDATA[cdata &amp;]]></t></si>"
                    + "<si><t>last</t></si>");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            for (int i = 0; i < expected.Length; i++)
            {
                Assert.Equal(expected[i], e.Current[i].GetString());
            }
        }

        [Fact]
        public void CommentContainingGreaterThanInsideSheetDataIsSkipped()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row><!-- a > b --><row r="2"><c r="A2"><v>2</v></c></row>""");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("1", e.Current[0].GetString());
            Assert.True(e.MoveNext());
            Assert.Equal("2", e.Current[0].GetString());
            Assert.False(e.MoveNext());
        }

        [Fact]
        public async Task AsyncCommentContainingGreaterThanInsideSheetDataIsSkipped()
        {
            await using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row><!-- a > b --><row r="2"><c r="A2"><v>2</v></c></row>""");
            await using XlsxReader reader = await Excel.FromXlsxAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator e = reader.GetAsyncEnumerator(TestContext.Current.CancellationToken);

            Assert.True(await e.MoveNextAsync());
            Assert.Equal("1", e.Current[0].GetString());
            Assert.True(await e.MoveNextAsync());
            Assert.Equal("2", e.Current[0].GetString());
            Assert.False(await e.MoveNextAsync());
        }

        [Fact]
        public void CDataInsideTextRunIsCopiedWithoutEntityDecoding()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="inlineStr"><is><t><![CDATA[raw &amp; <tag>]]></t></is></c></row>""");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("raw &amp; <tag>", e.Current[0].GetString());
        }

        [Fact]
        public void CDataInsideSharedStringIsCopiedWithoutEntityDecoding()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t><![CDATA[shared &amp; <tag>]]></t></si>");
            using XlsxReader reader = Excel.FromXlsx(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("shared &amp; <tag>", e.Current[0].GetString());
        }

        [Fact]
        public void ParseRelationshipsDoesNotMatchLongerElementNameSharingThePrefix()
        {
            var rels = XlsxXml.ParseRelationships(
                """<Relationship Id="rId1" Target="worksheets/sheet1.xml"/><RelationshipGroup Id="rId2" Target="malicious.xml"/>"""u8);

            Assert.Equal("worksheets/sheet1.xml", rels["rId1"]);
            Assert.False(rels.ContainsKey("rId2"));
        }
    }
}
