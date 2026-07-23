using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class XlsxXmlDialectTests
    {
        [Fact]
        public void SingleQuotedCellAttributesAreRead()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r='C1' t='s'><v>0</v></c></row>""",
                sharedStrings: "<si><t>single quoted</t></si>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("single quoted", e.Current[2].GetString());
        }

        [Fact]
        public void CommentContainingGreaterThanInsideSheetDataIsSkipped()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1"><v>1</v></c></row><!-- a > b --><row r="2"><c r="A2"><v>2</v></c></row>""");
            using XlsxReader reader = Excel.From(ms);
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
            await using XlsxReader reader = await Excel.FromAsync(ms, ct: TestContext.Current.CancellationToken);
            await using XlsxReader.Enumerator e = await reader.GetAsyncEnumeratorAsync(TestContext.Current.CancellationToken);

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
            using XlsxReader reader = Excel.From(ms);
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
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("shared &amp; <tag>", e.Current[0].GetString());
        }

        [Fact]
        public void ParseRelationshipsDoesNotMatchLongerElementNameSharingThePrefix()
        {
            // "<RelationshipGroup" is not a valid OPC element, but it contains "<Relationship" as a
            // literal substring — without a name-boundary check, TagSpanEnumerable would misparse it
            // as a real <Relationship> tag and inject a spurious rId into the map.
            var rels = XlsxXml.ParseRelationships(
                """<Relationship Id="rId1" Target="worksheets/sheet1.xml"/><RelationshipGroup Id="rId2" Target="malicious.xml"/>"""u8);

            Assert.Equal("worksheets/sheet1.xml", rels["rId1"]);
            Assert.False(rels.ContainsKey("rId2"));
        }
    }
}
