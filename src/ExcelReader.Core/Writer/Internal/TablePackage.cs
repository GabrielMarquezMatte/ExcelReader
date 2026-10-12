using System.Globalization;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class TablePackage
    {
        internal const string TableRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";

        internal static string RelationshipId(int index)
        {
            return string.Create(CultureInfo.InvariantCulture, $"rId{index + 1}");
        }

        internal static string SheetRelsXml(IReadOnlyList<WrittenTable> tables, string extension)
        {
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture, $"<Relationships xmlns=\"{XlsxConstants.PackageRelationshipsNs}\">");
            for (int i = 0; i < tables.Count; i++)
            {
                sb.Append(CultureInfo.InvariantCulture,
                    $"<Relationship Id=\"{RelationshipId(i)}\" Type=\"{TableRelType}\" Target=\"../tables/table{tables[i].Id}.{extension}\"/>");
            }
            sb.Append("</Relationships>");
            return sb.ToString();
        }
    }
}
