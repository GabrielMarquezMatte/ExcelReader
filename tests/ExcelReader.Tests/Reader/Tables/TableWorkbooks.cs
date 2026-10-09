using System.IO.Compression;
using System.Text;
using ExcelReader.Core.Reader.Xlsb;
using B = ExcelReader.Tests.Reader.Xlsb.Biff12Build;

namespace ExcelReader.Tests.Reader.Tables
{
    internal static class TableWorkbooks
    {
        internal const string TableRelationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
        private const string PackageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        private const string Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        internal static MemoryStream Xlsx(string rows, string tableXml, string target = "../tables/table1.xml", string extraRelationships = "")
        {
            string rels = $"""<Relationships xmlns="{PackageRelationships}">{extraRelationships}<Relationship Id="rId1" Type="{TableRelationship}" Target="{target}"/></Relationships>""";
            return WorkbookBuilder.BuildMultiSheet(
                [("S1", rows)],
                extraParts: [("xl/worksheets/_rels/sheet1.xml.rels", rels), ("xl/tables/table1.xml", tableXml)]);
        }

        internal static string Table(string reference, string columns, string attributes = "")
        {
            return $"""<table xmlns="{Main}" id="1" name="T" displayName="T" ref="{reference}"{attributes}><tableColumns>{columns}</tableColumns></table>""";
        }

        internal static MemoryStream Xlsb(byte[] tableBin)
        {
            MemoryStream ms = new();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                Add(zip, "xl/workbook.bin", B.Record(Brt.BundleSh, [.. B.U32(0), .. B.U32(0), .. B.WideString("rId1"), .. B.WideString("S1")]));
                Add(zip, "xl/_rels/workbook.bin.rels", Encoding.UTF8.GetBytes(
                    """<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.bin"/></Relationships>"""));
                Add(zip, "xl/styles.bin", []);
                Add(zip, "xl/sharedStrings.bin", []);
                Add(zip, "xl/worksheets/sheet1.bin",
                [
                    .. B.Record(Brt.RowHdr, B.U32(0)),
                    .. B.Record(Brt.CellSt, B.CellSt(0, 0, "H")),
                    .. B.Record(Brt.RowHdr, B.U32(1)),
                    .. B.Record(Brt.CellRk, B.CellRk(0, 0, (1u << 2) | 0x02)),
                    .. B.Record(Brt.EndSheetData),
                ]);
                Add(zip, "xl/worksheets/_rels/sheet1.bin.rels", Encoding.UTF8.GetBytes(
                    $"""<Relationships><Relationship Id="rId1" Type="{TableRelationship}" Target="../tables/table1.bin"/></Relationships>"""));
                Add(zip, "xl/tables/table1.bin", tableBin);
            }
            ms.Position = 0;
            return ms;
        }

        internal static byte[] BeginList(uint firstRow, uint lastRow, uint firstColumn, uint lastColumn, uint headerRowCount, uint totalsRowCount, string name)
        {
            return B.Record(Brt.BeginList,
            [
                .. B.U32(firstRow), .. B.U32(lastRow), .. B.U32(firstColumn), .. B.U32(lastColumn),
                .. B.U32(0), .. B.U32(1), .. B.U32(headerRowCount), .. B.U32(totalsRowCount), .. B.U32(0),
                .. Ones(24), .. B.U32(0),
                .. B.WideString(name), .. B.WideString(name), .. B.WideString(""),
                .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
            ]);
        }

        internal static byte[] ListColumn(uint id, string caption)
        {
            return B.Record(Brt.BeginListCol,
            [
                .. B.U32(id), .. B.U32(0), .. Ones(12), .. B.U32(0),
                .. B.NullWideString(), .. B.WideString(caption),
                .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(), .. B.NullWideString(),
            ]);
        }

        internal static byte[] StyleClient(string style)
        {
            return B.Record(Brt.ListTableStyleClient, [.. B.U16(4), .. B.WideString(style)]);
        }

        private static byte[] Ones(int count)
        {
            return Enumerable.Repeat((byte)0xFF, count).ToArray();
        }

        private static void Add(ZipArchive zip, string name, byte[] bytes)
        {
            using Stream stream = zip.CreateEntry(name).Open();
            stream.Write(bytes);
        }
    }
}
