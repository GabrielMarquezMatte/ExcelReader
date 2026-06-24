using System.Globalization;

namespace ExcelReader.Core.Writer.Internal
{
    internal static class CellFormatter
    {
        private static string BuildRef(int columnIndex, int rowNumber)
        {
            Span<char> colBuf = stackalloc char[3];
            int colLen = ColumnName.Write(colBuf, columnIndex);
            Span<char> rowBuf = stackalloc char[7];
            rowNumber.TryFormat(rowBuf, out int rowLen, default, CultureInfo.InvariantCulture);
            return string.Concat(colBuf[..colLen], rowBuf[..rowLen]);
        }

        internal static void WriteEmpty(StreamWriter xml, int columnIndex, int rowNumber)
        {
            xml.Write("<c r=\"");
            xml.Write(BuildRef(columnIndex, rowNumber));
            xml.Write("\"/>");
        }

        internal static void WriteString(StreamWriter xml, string value, int columnIndex, int rowNumber)
        {
            xml.Write("<c r=\"");
            xml.Write(BuildRef(columnIndex, rowNumber));
            xml.Write("\" t=\"inlineStr\"><is><t>");
            WriteEscaped(xml, value);
            xml.Write("</t></is></c>");
        }

        internal static void WriteBool(StreamWriter xml, bool value, int columnIndex, int rowNumber)
        {
            xml.Write("<c r=\"");
            xml.Write(BuildRef(columnIndex, rowNumber));
            xml.Write("\" t=\"b\"><v>");
            xml.Write(value ? '1' : '0');
            xml.Write("</v></c>");
        }

        internal static void WriteDateTime(StreamWriter xml, DateTime value, int columnIndex, int rowNumber)
        {
            xml.Write("<c r=\"");
            xml.Write(BuildRef(columnIndex, rowNumber));
            xml.Write("\" s=\"1\"><v>");
            double oaDate = value.ToOADate();
            Span<char> buf = stackalloc char[32];
            oaDate.TryFormat(buf, out int written, "G17", CultureInfo.InvariantCulture);
            xml.Write(buf[..written]);
            xml.Write("</v></c>");
        }

        internal static void WriteNumber<T>(StreamWriter xml, T value, int columnIndex, int rowNumber)
            where T : ISpanFormattable
        {
            xml.Write("<c r=\"");
            xml.Write(BuildRef(columnIndex, rowNumber));
            xml.Write("\"><v>");
            Span<char> buf = stackalloc char[64];
            value.TryFormat(buf, out int written, default, CultureInfo.InvariantCulture);
            xml.Write(buf[..written]);
            xml.Write("</v></c>");
        }

        private static void WriteEscaped(StreamWriter xml, string value)
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
                    xml.Write(value.AsSpan(start, i - start));
                }
                xml.Write(escape);
                start = i + 1;
            }
            if (start < value.Length)
            {
                xml.Write(value.AsSpan(start));
            }
        }
    }
}
