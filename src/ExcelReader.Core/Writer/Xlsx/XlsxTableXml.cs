using System.Globalization;
using System.Text;
using ExcelReader.Core.Writer.Internal;

namespace ExcelReader.Core.Writer.Xlsx
{
    internal static class XlsxTableXml
    {
        internal static string TableParts(int count)
        {
            StringBuilder sb = new();
            sb.Append(CultureInfo.InvariantCulture, $"<tableParts count=\"{count}\">");
            for (int i = 0; i < count; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<tablePart r:id=\"{TablePackage.RelationshipId(i)}\"/>");
            }
            sb.Append("</tableParts>");
            return sb.ToString();
        }

        internal static string Table(WrittenTable table)
        {
            string name = Escape(table.Name);
            string reference = table.Ref;
            StringBuilder sb = new();
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            sb.Append(CultureInfo.InvariantCulture,
                $"<table xmlns=\"{XlsxConstants.MainNs}\" id=\"{table.Id}\" name=\"{name}\" displayName=\"{name}\" ref=\"{reference}\" totalsRowShown=\"0\">");
            sb.Append(CultureInfo.InvariantCulture, $"<autoFilter ref=\"{reference}\"/>");
            sb.Append(CultureInfo.InvariantCulture, $"<tableColumns count=\"{table.Columns.Length}\">");
            for (int i = 0; i < table.Columns.Length; i++)
            {
                sb.Append(CultureInfo.InvariantCulture, $"<tableColumn id=\"{i + 1}\" name=\"{Escape(table.Columns[i])}\"/>");
            }
            sb.Append("</tableColumns>");
            AppendStyleInfo(sb, table.Options);
            sb.Append("</table>");
            return sb.ToString();
        }

        private static string Escape(string value)
        {
            using BiffBuffer buffer = new(value.Length + 16);
            CellFormatter.WriteEscaped(buffer, value);
            return Encoding.UTF8.GetString(buffer.Span);
        }

        private static void AppendStyleInfo(StringBuilder sb, ExcelTableOptions options)
        {
            sb.Append("<tableStyleInfo");
            if (options.StyleName is not null)
            {
                sb.Append(CultureInfo.InvariantCulture, $" name=\"{options.StyleName}\"");
            }
            sb.Append(CultureInfo.InvariantCulture,
                $" showFirstColumn=\"{Flag(options.ShowFirstColumn)}\" showLastColumn=\"{Flag(options.ShowLastColumn)}\" showRowStripes=\"{Flag(options.ShowRowStripes)}\" showColumnStripes=\"{Flag(options.ShowColumnStripes)}\"/>");
        }

        private static char Flag(bool value)
        {
            return value ? '1' : '0';
        }
    }
}
