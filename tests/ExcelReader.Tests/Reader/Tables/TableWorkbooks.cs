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
    }
}
