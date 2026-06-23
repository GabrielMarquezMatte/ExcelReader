using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Enums;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class SampleTest
    {
        [Fact]
        public void ReadsSharedStringsStylesAndNumbers()
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "sample.xlsx");
            using var reader = Excel.FromFile(path);

            int r = 0;
            foreach (var row in reader)
            {
                if (r == 0)
                {
                    // Header row: all shared strings, resolved to text.
                    Assert.Equal("file", row[0].GetString());
                    Assert.Equal("changes", row[1].GetString());
                    Assert.Equal("lines_added", row[2].GetString());
                    Assert.Equal("lines_deleted", row[3].GetString());
                    Assert.Equal(CellType.ExcelString, row[0].Type);
                }
                else if (r == 1)
                {
                    // Styled shared string (s="1") + UTF-8 numeric parse.
                    Assert.Equal("global.json", row[0].GetString());
                    Assert.Equal(1, row[0].StyleIndex);
                    Assert.True(row[1].TryParse(null, out int n));
                    Assert.Equal(2, n);
                    Assert.Equal(CellType.Number, row[1].Type);
                }
                r++;
            }
            Assert.Equal(3, r);
        }

        [Fact]
        public void HandlesSparseCellsAndBufferGrowth()
        {
            // A row with a gap (no B), and an inline string far larger than the 64 KB scan buffer
            // to exercise compaction/grow and the cross-boundary </c> search.
            string big = new('x', 100_000);
            using var ms = BuildWorkbook(
                $"""<row r="1"><c r="A1"><v>10</v></c><c r="C1"><v>30</v></c></row>""" +
                $"""<row r="2"><c r="A2" t="inlineStr"><is><t>{big}</t></is></c></row>""");

            using var reader = Excel.From(ms);
            int r = 0;
            foreach (var row in reader)
            {
                if (r == 0)
                {
                    Assert.Equal(3, row.ColumnCount);
                    Assert.True(row[0].TryParse(null, out int a));
                    Assert.Equal(10, a);
                    Assert.Equal(CellType.Empty, row[1].Type); // the gap
                    Assert.True(row[2].TryParse(null, out int c));
                    Assert.Equal(30, c);
                }
                else if (r == 1)
                {
                    Assert.Equal(big.Length, row[0].GetString().Length);
                }
                r++;
            }
            Assert.Equal(2, r);
        }

        [Fact]
        public void DecodesXmlEntitiesInSharedStrings()
        {
            using var ms = BuildWorkbook(
                """<row r="1"><c r="A1" t="s"><v>0</v></c></row>""",
                sharedStrings: "<si><t>a &amp; b &lt;tag&gt; &#65;</t></si>");

            using var reader = Excel.From(ms);
            using var enumerator = reader.GetEnumerator();
            if (!enumerator.MoveNext())
            {
                Assert.Fail("Expected at least one row");
            }
            var row = enumerator.Current;
            Assert.Equal("a & b <tag> A", row[0].GetString());
        }

        private static MemoryStream BuildWorkbook(string sheetRows, string? sharedStrings = null)
        {
            const string main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            const string pkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Write(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{main}" xmlns:r="{rel}"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{pkgRel}"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""");
                Write(zip, "xl/worksheets/sheet1.xml",
                    $"""<worksheet xmlns="{main}"><sheetData>{sheetRows}</sheetData></worksheet>""");
                if (sharedStrings is not null)
                {
                    Write(zip, "xl/sharedStrings.xml", $"""<sst xmlns="{main}">{sharedStrings}</sst>""");
                }
            }
            ms.Position = 0;
            return ms;

            static void Write(ZipArchive zip, string name, string content)
            {
                using var s = zip.CreateEntry(name).Open();
                var bytes = Encoding.UTF8.GetBytes(content);
                s.Write(bytes, 0, bytes.Length);
            }
        }
    }
}
