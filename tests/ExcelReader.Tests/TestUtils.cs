using System.IO.Compression;
using System.Text;

namespace ExcelReader.Tests
{
    internal static class WorkbookBuilder
    {
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private const string PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        internal static MemoryStream Build(string sheetRows, string? sharedStrings = null, string? styles = null, bool date1904 = false)
        {
            return BuildMultiSheet([("S1", sheetRows)], sharedStrings, styles, date1904);
        }

        internal static MemoryStream BuildMultiSheet(
            (string Name, string Rows)[] sheets,
            string? sharedStrings = null,
            string? styles = null,
            bool date1904 = false)
        {
            var sheetXml = new string[sheets.Length];
            var relXml = new string[sheets.Length];
            var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                for (int i = 0; i < sheets.Length; i++)
                {
                    int id = i + 1;
                    var (name, rows) = sheets[i];
                    sheetXml[i] = $"""<sheet name="{name}" sheetId="{id}" r:id="rId{id}"/>""";
                    relXml[i] = $"""<Relationship Id="rId{id}" Type="x" Target="worksheets/sheet{id}.xml"/>""";
                    Write(zip, $"xl/worksheets/sheet{id}.xml",
                        $"""<worksheet xmlns="{Main}"><sheetData>{rows}</sheetData></worksheet>""");
                }
                string workbookPr = date1904 ? """<workbookPr date1904="1"/>""" : "";
                Write(zip, "xl/workbook.xml",
                    $"""<workbook xmlns="{Main}" xmlns:r="{Rel}">{workbookPr}<sheets>{string.Concat(sheetXml)}</sheets></workbook>""");
                Write(zip, "xl/_rels/workbook.xml.rels",
                    $"""<Relationships xmlns="{PkgRel}">{string.Concat(relXml)}</Relationships>""");
                if (sharedStrings is not null)
                {
                    Write(zip, "xl/sharedStrings.xml", $"""<sst xmlns="{Main}">{sharedStrings}</sst>""");
                }
                if (styles is not null)
                {
                    string withNs = styles.Replace("<styleSheet>",
                        $"""<styleSheet xmlns="{Main}">""", StringComparison.Ordinal);
                    Write(zip, "xl/styles.xml", $"""<?xml version="1.0"?>{withNs}""");
                }
            }
            ms.Position = 0;
            return ms;
        }

        private static void Write(ZipArchive zip, string name, string content)
        {
            using var s = zip.CreateEntry(name).Open();
            byte[] bytes = Encoding.UTF8.GetBytes(content);
            s.Write(bytes, 0, bytes.Length);
        }
    }
}
