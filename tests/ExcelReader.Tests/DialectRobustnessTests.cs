using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;
using ExcelReader.Core.ValueObjects;
using B = ExcelReader.Tests.Biff12Build;

namespace ExcelReader.Tests
{
    public class DialectRobustnessTests
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";


        [Fact]
        public void AttrReadsDoubleQuotedValue()
        {
            ReadOnlySpan<byte> tag = " name=\"S1\" r:id=\"rId1\""u8;
            Assert.True(XlsxXml.Attr(tag, " name="u8).SequenceEqual("S1"u8));
            Assert.True(XlsxXml.Attr(tag, " r:id="u8).SequenceEqual("rId1"u8));
        }

        [Fact]
        public void AttrReadsSingleQuotedValue()
        {
            ReadOnlySpan<byte> tag = " name='S1' r:id='rId1'"u8;
            Assert.True(XlsxXml.Attr(tag, " name="u8).SequenceEqual("S1"u8));
            Assert.True(XlsxXml.Attr(tag, " r:id="u8).SequenceEqual("rId1"u8));
        }

        [Fact]
        public void AttrReturnsEmptyForMissingOrUnquotedAttribute()
        {
            Assert.True(XlsxXml.Attr(" name=\"S1\""u8, " id="u8).IsEmpty);
            Assert.True(XlsxXml.Attr(" name=bad"u8, " name="u8).IsEmpty);
        }


        [Fact]
        public void OverlongDecimalEntityRoundTripsAsLiteralText()
        {
            string result = XlsxXml.DecodeToString("&#99999999999999999999;"u8);
            Assert.Equal("&#99999999999999999999;", result);
        }

        [Fact]
        public void OverlongHexEntityRoundTripsAsLiteralText()
        {
            string result = XlsxXml.DecodeToString("&#xFFFFFFFFFFFFFFFF;"u8);
            Assert.Equal("&#xFFFFFFFFFFFFFFFF;", result);
        }

        [Fact]
        public void ValidNumericEntityStillDecodesNormally()
        {
            Assert.Equal("A", XlsxXml.DecodeToString("&#65;"u8));
            Assert.Equal("A", XlsxXml.DecodeToString("&#x41;"u8));
        }

        [Fact]
        public void SingleQuotedWorkbookAttributesAreRead()
        {
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                WriteEntry(zip, "xl/worksheets/sheet1.xml",
                    $"""<worksheet xmlns="{Main}"><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>""");
                WriteEntry(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{Main}" xmlns:r="{Rel}"><workbookPr date1904='1'/><sheets><sheet name='Data' sheetId='1' r:id='rId1'/></sheets></workbook>""");
                WriteEntry(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}"><Relationship Id='rId1' Type='x' Target='worksheets/sheet1.xml'/></Relationships>""");
            }
            ms.Position = 0;

            using XlsxReader reader = Excel.From(ms);
            Assert.Equal("Data", reader.SheetName);
            Assert.True(reader.IsDate1904);
            using XlsxReader.Enumerator e = reader.GetEnumerator();
            Assert.True(e.MoveNext());
            Assert.Equal("1", e.Current[0].GetString());
        }

        [Fact]
        public void SingleQuotedStyleAttributesDetectDate()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" s="0"><v>45658</v></c></row>""",
                styles: "<styleSheet><numFmts count='1'><numFmt numFmtId='164' formatCode='yyyy-mm-dd'/></numFmts><cellXfs count='1'><xf numFmtId='164'/></cellXfs></styleSheet>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
        }


        [Fact]
        public void CorruptSharedStringIndexYieldsEmptyNotFirstString()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v>x</v></c><c r="B1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>first</t></si><si><t>second</t></si>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("", e.Current[0].GetString());
            Assert.Equal("first", e.Current[1].GetString());
        }

        [Fact]
        public void EmptySharedStringIndexYieldsEmpty()
        {
            using MemoryStream ms = WorkbookBuilder.Build(
                """<row r="1"><c r="A1" t="s"><v></v></c></row>""",
                sharedStrings: "<si><t>first</t></si>");
            using XlsxReader reader = Excel.From(ms);
            using XlsxReader.Enumerator e = reader.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("", e.Current[0].GetString());
        }


        [Fact]
        public void TruncatedXlsbStreamThrows()
        {
            byte[] sheet = TruncatedSheet();
            using XlsbReader reader = new(sharedFlat: [], sharedOffsets: [0], styleIsDate: [], date1904: false);
            using XlsbReader.Enumerator e = new(reader, new MemoryStream(sheet));

            Assert.Throws<InvalidDataException>(() =>
            {
                while (e.MoveNext())
                {
                }
            });
        }

        [Fact]
        public async Task TruncatedXlsbStreamThrowsAsync()
        {
            byte[] sheet = TruncatedSheet();
            await using XlsbReader reader = new(sharedFlat: [], sharedOffsets: [0], styleIsDate: [], date1904: false);
            await using XlsbReader.Enumerator e = new(reader, new MemoryStream(sheet));

            await Assert.ThrowsAsync<InvalidDataException>(async () =>
            {
                while (await e.MoveNextAsync())
                {
                }
            });
        }

        [Fact]
        public void CompleteXlsbStreamWithoutEndMarkerDoesNotThrow()
        {
            byte[] sheet =
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellReal, B.CellReal(0, 0, 3.14)),
            ];
            using XlsbReader reader = new(sharedFlat: [], sharedOffsets: [0], styleIsDate: [], date1904: false);
            using XlsbReader.Enumerator e = new(reader, new MemoryStream(sheet));

            Assert.True(e.MoveNext());
            Assert.Equal(3.14, Read(e.Current[0]));
            Assert.False(e.MoveNext());
        }

        private static byte[] TruncatedSheet()
        {
            byte[] cell = B.Record(Brt.CellReal, B.CellReal(1, 0, 2.0));
            return
            [
                .. B.Record(Brt.RowHdr),
                .. B.Record(Brt.CellReal, B.CellReal(0, 0, 3.14)),
                .. cell[..3],
            ];
        }

        private static double Read(Cell cell)
        {
            Assert.True(cell.TryGetDouble(out double value));
            return value;
        }

        private static void WriteEntry(ZipArchive zip, string name, string content)
        {
            using Stream s = zip.CreateEntry(name).Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            s.Write(bytes, 0, bytes.Length);
        }
    }
}
