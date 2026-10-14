using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests.Reader
{
    /// <summary>
    /// Text as Excel itself stores it. <c>excel-formula-strings</c> was saved by Excel 16 in both formats:
    /// A1 holds a line break, B1 and C1 formulas whose text has whitespace at an end, B2:B4 a formula
    /// filled down, and C2:C4 an array formula.
    /// </summary>
    public sealed class ExcelWrittenTextTests
    {
        private static readonly string[][] Expected =
        [
            ["a\nb", "\nnl", " pad "],
            ["1", "1s", "1a"],
            ["2", "2s", "2a"],
            ["3", "3s", "3a"],
        ];

        [Theory]
        [InlineData("xlsx")]
        [InlineData("xls")]
        public void AnExcelFileReadsAsExcelShowsIt(string extension)
        {
            string path = Path.Combine(AppContext.BaseDirectory, "data", "excel-formula-strings." + extension);

            using IExcelWorkbook workbook = Excel.Open(path);

            Assert.Equal(Expected, Rows(workbook, 3));
        }

        [Fact]
        public void ALiteralCarriageReturnIsAnXmlLineEnd()
        {
            byte[] package = Xlsx(
                "<row r=\"1\">" +
                "<c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c><c r=\"C1\" t=\"s\"><v>2</v></c>" +
                "<c r=\"D1\" t=\"inlineStr\"><is><t>in\r\nline</t></is></c>" +
                "<c r=\"E1\" t=\"str\"><f>X</f><v>for\r\nmula</v></c>" +
                "<c r=\"F1\" t=\"s\"><v>3</v></c><c r=\"G1\" t=\"s\"><v>4</v></c>" +
                "</row>",
                "<si><t>a\r\nb</t></si>" +
                "<si><t>lone\rcr</t></si>" +
                "<si><t>kept&#13;\r\nref and _x000D_\r\nescape</t></si>" +
                "<si><r><t>rich\r\n</t></r><r><t>run</t></r></si>" +
                "<si><t>a_b\r\nc</t></si>");

            using IExcelWorkbook workbook = Excel.Open(package);

            string[] row = Assert.Single(Rows(workbook, 7));
            Assert.Equal(
                ["a\nb", "lone\ncr", "kept\r\nref and \r\nescape", "in\nline", "for\nmula", "rich\nrun", "a_b\nc"],
                row);
        }

        [Fact]
        public void AValueElementMayCarryAttributes()
        {
            byte[] package = Xlsx(
                "<row r=\"1\">" +
                "<c r=\"A1\" t=\"str\"><f>X</f><v xml:space=\"preserve\"> padded </v></c>" +
                "<c r=\"B1\"><v xml:space=\"preserve\">42</v></c>" +
                "<c r=\"C1\" t=\"str\"><f>X</f><v>plain</v></c>" +
                "<c r=\"D1\" t=\"str\"><f>X</f><v/></c>" +
                "<c r=\"E1\" t=\"s\"><v xml:space=\"preserve\">0</v></c>" +
                "</row>",
                "<si><t>shared</t></si>");

            using IExcelWorkbook workbook = Excel.Open(package);

            string[] row = Assert.Single(Rows(workbook, 5));
            Assert.Equal([" padded ", "42", "plain", "", "shared"], row);
        }

        private static List<string[]> Rows(IExcelWorkbook workbook, int columns)
        {
            var rows = new List<string[]>();
            using IExcelRowEnumerator enumerator = workbook.FirstSheet.GetEnumerator();
            while (enumerator.MoveNext())
            {
                string[] cells = new string[columns];
                for (int c = 0; c < columns; c++)
                {
                    cells[c] = enumerator.Current[c].GetString();
                }
                rows.Add(cells);
            }
            return rows;
        }

        private static byte[] Xlsx(string rows, string sharedItems)
        {
            const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            using MemoryStream ms = new();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Add(zip, "[Content_Types].xml",
                    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
                    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
                    "<Default Extension=\"xml\" ContentType=\"application/xml\"/></Types>");
                Add(zip, "_rels/.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{Rel}/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
                Add(zip, "xl/workbook.xml",
                    $"<workbook xmlns=\"{Main}\" xmlns:r=\"{Rel}\"><sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                Add(zip, "xl/_rels/workbook.xml.rels",
                    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                    $"<Relationship Id=\"rId1\" Type=\"{Rel}/worksheet\" Target=\"worksheets/sheet1.xml\"/>" +
                    $"<Relationship Id=\"rId2\" Type=\"{Rel}/sharedStrings\" Target=\"sharedStrings.xml\"/></Relationships>");
                Add(zip, "xl/sharedStrings.xml", $"<sst xmlns=\"{Main}\">{sharedItems}</sst>");
                Add(zip, "xl/worksheets/sheet1.xml", $"<worksheet xmlns=\"{Main}\"><sheetData>{rows}</sheetData></worksheet>");
            }
            return ms.ToArray();
        }

        private static void Add(ZipArchive zip, string name, string content)
        {
            using Stream entry = zip.CreateEntry(name).Open();
            entry.Write(Encoding.UTF8.GetBytes(content));
        }
    }
}
