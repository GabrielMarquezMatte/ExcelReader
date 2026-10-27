using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Xlsx;

namespace ExcelReader.Tests.Reader.Xlsx
{
    public class XlsxLoadingRegressionTests
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        private static byte[] Build(string rows, string? sharedStrings = null, string? styles = null, string workbook = "", bool corePropsFirst = false)
        {
            string wb = workbook.Length > 0 ? workbook
                : $"""<workbook xmlns="{Main}" xmlns:r="{Rel}"><sheets><sheet name="S1" sheetId="1" r:id="rId1"/></sheets></workbook>""";
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                if (corePropsFirst)
                {
                    Write(zip, "docProps/core.xml", """<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"><dc:title xmlns:dc="http://purl.org/dc/elements/1.1/">t</dc:title></cp:coreProperties>""");
                }
                Write(zip, "xl/workbook.xml", wb);
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}"><Relationship Id="rId1" Type="x" Target="worksheets/sheet1.xml"/></Relationships>""");
                Write(zip, "xl/worksheets/sheet1.xml", $"""<worksheet xmlns="{Main}"><sheetData>{rows}</sheetData></worksheet>""");
                if (sharedStrings is not null)
                {
                    Write(zip, "xl/sharedStrings.xml", sharedStrings);
                }
                if (styles is not null)
                {
                    Write(zip, "xl/styles.xml", styles);
                }
            }
            return ms.ToArray();
        }

        private static void Write(ZipArchive zip, string path, string content)
        {
            using Stream entry = zip.CreateEntry(path, CompressionLevel.NoCompression).Open();
            entry.Write(Encoding.UTF8.GetBytes(content));
        }

        [Fact]
        public void Should_ReadSharedStrings_When_StoredInMemoryAfterPrefixedCoreProps()
        {
            byte[] xlsx = Build("""<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>""",
                $"""<sst xmlns="{Main}"><si><t>alpha</t></si><si><t>beta</t></si></sst>""", corePropsFirst: true);

            using XlsxWorkbook reader = Excel.FromXlsx(xlsx.AsMemory());
            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("alpha", e.Current[0].GetString());
            Assert.Equal("beta", e.Current[1].GetString());
        }

        [Fact]
        public void Should_LoadWorkbook_When_UniqueCountExceedsActualStrings()
        {
            byte[] xlsx = Build("""<row r="1"><c r="A1" t="s"><v>1</v></c></row>""",
                $"""<sst xmlns="{Main}" count="5000" uniqueCount="5000"><si><t>a</t></si><si><t>b</t></si></sst>""");

            using XlsxWorkbook reader = Excel.FromXlsx(xlsx.AsMemory());
            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("b", e.Current[0].GetString());
        }

        [Fact]
        public void Should_ResolveRelationship_When_IdUsesNonRPrefix()
        {
            string wb = $"""<workbook xmlns="{Main}" xmlns:ns1="{Rel}"><sheets><sheet name="S1" sheetId="1" ns1:id="rId1"/></sheets></workbook>""";
            byte[] xlsx = Build("""<row r="1"><c r="A1"><v>7</v></c></row>""", workbook: wb);

            using XlsxWorkbook reader = Excel.FromXlsx(xlsx.AsMemory());
            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal("7", e.Current[0].GetString());
        }

        [Fact]
        public void Should_IgnoreDxfNumFmts_When_ClassifyingCellStyles()
        {
            string styles = $"""<styleSheet xmlns="{Main}"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="1"><xf numFmtId="164"/></cellXfs><dxfs count="1"><dxf><numFmt numFmtId="164" formatCode="0.00"/></dxf></dxfs></styleSheet>""";
            byte[] xlsx = Build("""<row r="1"><c r="A1" s="0"><v>45000</v></c></row>""", styles: styles);

            using XlsxWorkbook reader = Excel.FromXlsx(xlsx.AsMemory());
            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
        }

        [Fact]
        public void Should_ReturnTimeFraction_When_DateCellHoldsTimeOnly()
        {
            byte[] xlsx = Build("""<row r="1"><c r="A1" t="d"><v>12:30:00</v></c></row>""");

            using XlsxWorkbook reader = Excel.FromXlsx(xlsx.AsMemory());
            using XlsxWorkbook.Enumerator e = reader.FirstSheet.GetEnumerator();

            Assert.True(e.MoveNext());
            Assert.Equal(CellType.Date, e.Current[0].Type);
            double serial = double.Parse(e.Current[0].GetString(), System.Globalization.CultureInfo.InvariantCulture);
            Assert.Equal(12.5 / 24, serial, 9);
        }
    }
}
