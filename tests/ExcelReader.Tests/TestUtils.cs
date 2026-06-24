using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Writer;

namespace ExcelReader.Tests
{
    // Marks a skipped (empty) column gap inside a TypedWorkbook row.
    // A record class (not struct) so `new Gap()` honors the Count = 1 default.
    internal sealed record Gap(int Count = 1);

    // Builds workbooks via the real WorkbookWriter from typed cell values.
    // Use for reader/parser fixtures expressible as inline strings, numbers,
    // dates (builtin numFmt 14), and bools. For shared strings, custom number
    // formats, the 1904 date system, or error/formula cells, use WorkbookBuilder
    // (raw XML) instead — WorkbookWriter cannot emit those.
    internal static class TypedWorkbook
    {
        // Single sheet "S1"; each row is an array of cell values.
        internal static Task<MemoryStream> BuildAsync(params object?[][] rows)
        {
            return BuildMultiSheetAsync(("S1", rows));
        }

        internal static async Task<MemoryStream> BuildMultiSheetAsync(
            params (string Name, object?[][] Rows)[] sheets)
        {
            var ms = new MemoryStream();
            await using (WorkbookWriter wb = await WorkbookWriter.CreateAsync(ms, leaveOpen: true))
            {
                await wb.StartAsync();
                foreach ((string name, object?[][] rows) in sheets)
                {
                    SheetWriter sheet = wb.AddSheet(name);
                    await sheet.StartAsync();
                    foreach (object?[] row in rows)
                    {
                        await using RowWriter rw = await sheet.StartRowAsync();
                        foreach (object? cell in row)
                        {
                            WriteCell(rw, cell);
                        }
                    }
                    await sheet.EndAsync();
                }
                await wb.EndAsync();
            }
            ms.Position = 0;
            return ms;
        }

        private static void WriteCell(RowWriter rw, object? cell)
        {
            switch (cell)
            {
                case null: rw.Write((string?)null); break;
                case string s: rw.Write(s); break;
                case bool b: rw.Write(b); break;
                case int i: rw.Write(i); break;
                case long l: rw.Write(l); break;
                case double d: rw.Write(d); break;
                case decimal m: rw.Write(m); break;
                case DateTime dt: rw.Write(dt); break;
                case Gap g: rw.Skip(g.Count); break;
                default: throw new NotSupportedException($"Unsupported cell value type: {cell.GetType()}");
            }
        }
    }

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
