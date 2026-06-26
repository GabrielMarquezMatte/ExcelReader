using System.Globalization;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        // Writes the cell reference (e.g. "B7") directly to the writer.
        // Max XLSX cell is XFD1048576 -> 3 column letters + 7 row digits.
        private static void WriteRef(StreamWriter xml, int columnIndex, int rowNumber)
        {
            Span<char> buf = stackalloc char[10];
            int len = ColumnName.Write(buf, columnIndex);
            rowNumber.TryFormat(buf[len..], out int rowLen, default, CultureInfo.InvariantCulture);
            xml.Write(buf[..(len + rowLen)]);
        }

        private static void WriteCellOpen(StreamWriter xml, int columnIndex, int rowNumber, bool includeReference)
        {
            xml.Write("<c");
            if (!includeReference)
            {
                return;
            }
            xml.Write(" r=\"");
            WriteRef(xml, columnIndex, rowNumber);
            xml.Write('"');
        }

        internal static void WriteEmpty(StreamWriter xml, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("/>");
        }

        internal static void WriteString(StreamWriter xml, string value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" t=\"inlineStr\"><is><t>");
            WriteEscaped(xml, value);
            xml.Write("</t></is></c>");
        }

        internal static void WriteBool(StreamWriter xml, bool value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" t=\"b\"><v>");
            xml.Write(value ? '1' : '0');
            xml.Write("</v></c>");
        }

        internal static void WriteDateTime(StreamWriter xml, DateTime value, int columnIndex, int rowNumber, bool includeReference)
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write(" s=\"1\"><v>");
            double oaDate = value.ToOADate();
            Span<char> buf = stackalloc char[32];
            oaDate.TryFormat(buf, out int written, "G17", CultureInfo.InvariantCulture);
            xml.Write(buf[..written]);
            xml.Write("</v></c>");
        }

        internal static void WriteNumber<T>(StreamWriter xml, T value, int columnIndex, int rowNumber, bool includeReference)
            where T : ISpanFormattable
        {
            WriteCellOpen(xml, columnIndex, rowNumber, includeReference);
            xml.Write("><v>");
            Span<char> buf = stackalloc char[64];
            value.TryFormat(buf, out int written, default, CultureInfo.InvariantCulture);
            xml.Write(buf[..written]);
            xml.Write("</v></c>");
        }

        private static void WriteEscaped(StreamWriter xml, ReadOnlySpan<char> value)
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
                    xml.Write(value[start..i]);
                }
                xml.Write(escape);
                start = i + 1;
            }
            if (start < value.Length)
            {
                xml.Write(value[start..]);
            }
        }
    }
}
