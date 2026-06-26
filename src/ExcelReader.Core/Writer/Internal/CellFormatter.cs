using System.Globalization;
using System.Text;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        // Writes the cell reference (e.g. "B7") directly to the writer.
        // Max XLSX cell is XFD1048576 -> 3 column letters + 7 row digits.
        private static void WriteRef(StringBuilder xml, int columnIndex, int rowNumber)
        {
            Span<char> buf = stackalloc char[10];
            int len = ColumnName.Write(buf, columnIndex);
            rowNumber.TryFormat(buf[len..], out int rowLen, default, CultureInfo.InvariantCulture);
            xml.Append(buf[..(len + rowLen)]);
        }

        private static void WriteCellOpen(StringBuilder xml, int columnIndex, int rowNumber, bool includeReference)
        {
            xml.Append("<c");
            if (!includeReference)
            {
                return;
            }
            xml.Append(" r=\"");
            WriteRef(xml, columnIndex, rowNumber);
            xml.Append('"');
        }

        internal static void WriteEmpty(StringBuilder xml, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Append("/>");
        }

        internal static void WriteString(StringBuilder xml, string value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Append(" t=\"inlineStr\"><is><t>");
            WriteEscaped(xml, value);
            xml.Append("</t></is></c>");
        }

        internal static void WriteBool(StringBuilder xml, bool value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Append(" t=\"b\"><v>");
            xml.Append(value ? '1' : '0');
            xml.Append("</v></c>");
        }

        internal static void WriteDateTime(StringBuilder xml, DateTime value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Append(" s=\"1\"><v>");
            double oaDate = value.ToOADate();
            Span<char> buf = stackalloc char[32];
            oaDate.TryFormat(buf, out int written, "G17", CultureInfo.InvariantCulture);
            xml.Append(buf[..written]);
            xml.Append("</v></c>");
        }

        internal static void WriteNumber<T>(StringBuilder xml, T value, int columnIndex, int rowNumber, bool includeReference)
            where T : ISpanFormattable
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Append("><v>");
            Span<char> buf = stackalloc char[64];
            value.TryFormat(buf, out int written, default, CultureInfo.InvariantCulture);
            xml.Append(buf[..written]);
            xml.Append("</v></c>");
        }

        private static void WriteEscaped(StringBuilder xml, ReadOnlySpan<char> value)
        {
            int start = 0;
            for (int i = 0; i < value.Length; i++)
            {
                string? escape = value[i] switch
                {
                    '&' => "&amp;",
                    '<' => "&lt;",
                    '>' => "&gt;",
                    '"' => "&quot;",
                    '\'' => "&apos;",
                    _ => null,
                };
                if (escape is null)
                {
                    continue;
                }
                if (i > start)
                {
                    xml.Append(value[start..i]);
                }
                xml.Append(escape);
                start = i + 1;
            }
            if (start < value.Length)
            {
                xml.Append(value[start..]);
            }
        }
    }
}
